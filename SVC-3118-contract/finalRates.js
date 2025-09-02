const XLSX = require('xlsx');
const fs = require('fs');
// Note: `node-fetch` may trigger a `punycode` deprecation warning in Node.js 18+.
// Consider using `undici` (Node.js built-in fetch) as an alternative if needed.
const fetch = require('node-fetch');
const { login } = require('./auth.js');

async function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Read Google Sheet as CSV
async function loadSheetData() {
    const SPREADSHEET_ID = '1yBg3JcGlt_Jhnegd-AOEAx83m6xSm5Ee3afZWcMBSNI';
    const GID = '1057237072';
    const CSV_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=csv&gid=${GID}`;

    try {
        const response = await fetch(CSV_URL, { timeout: 10000 });
        if (!response.ok) {
            throw new Error(`Failed to fetch CSV: HTTP ${response.status}`);
        }
        const csvText = await response.text();
        const rows = csvText.split('\n').map(row => 
            row.split(',').map(cell => cell.trim().replace(/\r$/, ''))
        );

        if (rows.length === 0) {
            return [];
        }
        return rows;
    } catch (error) {
        throw error;
    }
}

async function findServiceIdAndName(polName, podName, sheetData) {
    if (!polName || !podName) return { serviceId: '', serviceName: '' };
    const polSearch = polName.trim().toLowerCase();
    const podSearch = podName.trim().toLowerCase();

    const headers = sheetData[0] || [];
    const loadPortIdx = headers.indexOf('Load Port');
    const dischargePortIdx = headers.indexOf('Discharge Port');
    const serviceNameIdx = headers.indexOf('Service Name');
    const serviceIdIdx = headers.indexOf('Service ID');
    if (loadPortIdx === -1 || dischargePortIdx === -1) {
        return { serviceId: '', serviceName: '' };
    }

    for (let i = 1; i < sheetData.length; i++) {
        const row = sheetData[i];
        let loadPorts = row[loadPortIdx] ? row[loadPortIdx].trim().toLowerCase() : '';
        const dischargePort = row[dischargePortIdx] ? row[dischargePortIdx].trim().toLowerCase() : '';
        loadPorts = loadPorts.split(',').map(port => port.trim());
        if (loadPorts.includes(polSearch) && dischargePort === podSearch) {
            return {
                serviceId: row[serviceIdIdx] ? String(row[serviceIdIdx]) : '',
                serviceName: row[serviceNameIdx] || ''
            };
        }
    }
    return { serviceId: '', serviceName: '' };
}

async function sendServiceToMicroservice(polName, podName, serviceId, serviceName) {
    const data = [{
        origin_id: polName,
        discharge_id: podName,
        service_name: serviceName,
        service_id: serviceId
    }];
    try {
        const response = await fetch('https://hook.us1.make.com/3ewsb0bi54wrow4b0ivp695wrc2n8ijk', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data),
            timeout: 10000
        });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        console.log(`Successfully updated Google Sheet for ${polName} -> ${podName}`);
        return true;
    } catch (error) {
        console.error(`Failed to send service to microservice for ${polName} -> ${podName}: ${error.message}`);
        return false;
    }
}

async function sendRatesToEndpoint(rates, indices, sheetName) {
    try {
        const response = await fetch('https://hook.us1.make.com/pb3chu1cv36412wo9nzjeupbwo5ucvsw', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(rates),
            timeout: 10000
        });
        if (!response.ok) {
            console.error(`Failed to send ${rates.length} rates to endpoint for ${sheetName}: HTTP ${response.status}`);
            return false;
        }
        console.log(`Successfully sent ${rates.length} rates to endpoint for ${sheetName}`);
        return true;
    } catch (error) {
        console.error(`Error sending ${rates.length} rates to endpoint for ${sheetName}: ${error.message}`);
        return false;
    }
}

async function createFinalRates(finalRates, contractEffectiveDate, contractExpirationDate, sheetName = 'Unknown') {
    const timestamp = new Date().toISOString();

    // Initialize counters and tracking
    const missingFieldsCount = {
        port_destination: 0,
        port_origin: 0,
        port_discharge: 0,
        service: 0,
        port_origin_names: [],
        missing_service_routes: []
    };
    let sheetMatchCount = 0;
    let apiMatchCount = 0;
    let apiCallCount = 0;
    let apiFailedCount = 0;
    let apiSkippedCount = 0;
    let microserviceUpdateCount = 0;
    const maxApiCallsPerRun = 30;
    const maxRatesToSendPerRun = 100; 
    const missingServices = [];
    const newSheetEntries = [];
    const skippedApiRates = [];
    const unprocessedRates = [];

    // Load tracking state from file to persist across runs
    let trackingState;
    const trackingStateFile = 'trackingState.json';
    if (fs.existsSync(trackingStateFile)) {
        try {
            trackingState = JSON.parse(fs.readFileSync(trackingStateFile, 'utf8'));
            // Validate trackingState
            if (!trackingState.sentRateIndices || !Array.isArray(trackingState.sentRateIndices)) {
                console.warn(`Invalid trackingState in ${trackingStateFile}, resetting to empty.`);
                trackingState = { sentRateIndices: [], totalRatesSent: 0 };
            }
        } catch (error) {
            console.warn(`Failed to load ${trackingStateFile}: ${error.message}, resetting to empty.`);
            trackingState = { sentRateIndices: [], totalRatesSent: 0 };
        }
    } else {
        trackingState = { sentRateIndices: [], totalRatesSent: 0 };
    }

    // Authenticate with Betty Blocks API
    let jwtToken;
    try {
        const { jwtToken: token } = await login();
        jwtToken = token;
    } catch (error) {
        console.error(`Authentication failed for ${sheetName}: ${error.message}`);
        process.exit(1);
    }

    // Load CMA Parser Rates data
    let sheetData;
    try {
        sheetData = await loadSheetData();
    } catch (error) {
        console.error(`Failed to load sheet data for ${sheetName}: ${error.message}`);
        process.exit(1);
    }

    // Step 1: Check Google Sheet for service IDs
    const ratesToProcessViaApi = [];
    for (const rate of finalRates) {
        const { serviceId, serviceName } = await findServiceIdAndName(rate.polName, rate.podName, sheetData);
        if (serviceId) {
            sheetMatchCount++;
            rate.service = serviceId;
            rate.service_name = serviceName;
            newSheetEntries.push({
                'Load Port': rate.polName || '',
                'Discharge Port': rate.podName || '',
                'Delivery': rate['Place of Delivery'] ? String(rate['Place of Delivery']) : '',
                'Service Name': serviceName,
                'Service ID': serviceId
            });
        } else {
            ratesToProcessViaApi.push(rate);
        }
    }

    // Step 2: Process rates via Betty Blocks API in batches of 30
    if (ratesToProcessViaApi.length === 0) {
        console.log(`All service IDs found in Google Sheet for ${sheetName}, no API calls needed.`);
    } else {
        const ratesToFetch = ratesToProcessViaApi.slice(0, maxApiCallsPerRun);
        for (const rate of ratesToFetch) {
            // Validate POL and POD are numeric
            const polId = rate.POL ? parseInt(rate.POL) : null;
            const podId = rate.POD ? parseInt(rate.POD) : null;
            if (!polId || isNaN(polId) || !podId || isNaN(podId)) {
                apiSkippedCount++;
                skippedApiRates.push({
                    polName: rate.polName || 'N/A',
                    podName: rate.podName || 'N/A',
                    POL: rate.POL || 'N/A',
                    POD: rate.POD || 'N/A'
                });
                continue;
            }

            try {
                if (apiCallCount > 0) {
                    await delay(2000); // 2-second delay between requests
                }
                const response = await fetch('https://primefreight-development.betty.app/api/runtime/da93364a26fb4eeb9e56351ecec79abb', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${jwtToken}`
                    },
                    body: JSON.stringify({
                        query: "mutation($id: ID!, $input: ActionInput!) { action(id: $id, input: $input) }",
                        variables: {
                            id: "32c4095339884c2da3149b3a8c68bb11",
                            input: {
                                discharge_id: podId,
                                origin_id: polId
                            }
                        }
                    }),
                    timeout: 10000
                });

                const result = await response.json();
                apiCallCount++;
                if (result.data && result.data.action && result.data.action.results) {
                    const serviceId = result.data.action.results.service_id || '';
                    const serviceName = result.data.action.results.service_name || '';
                    if (serviceId) {
                        apiMatchCount++;
                        rate.service = serviceId;
                        rate.service_name = serviceName;
                        newSheetEntries.push({
                            'Load Port': rate.polName || '',
                            'Discharge Port': rate.podName || '',
                            'Delivery': rate['Place of Delivery'] ? String(rate['Place of Delivery']) : '',
                            'Service Name': serviceName,
                            'Service ID': serviceId
                        });
                        const success = await sendServiceToMicroservice(rate.polName, rate.podName, serviceId, serviceName);
                        if (success) {
                            microserviceUpdateCount++;
                        }
                    } else {
                        apiFailedCount++;
                    }
                } else {
                    apiFailedCount++;
                }
            } catch (error) {
                apiCallCount++;
                apiFailedCount++;
                console.error(`API call failed for ${rate.polName} -> ${rate.podName} in ${sheetName}: ${error.message}`);
            }
        }
    }

    // Step 3: Update missing services and track missing routes
    for (const rate of finalRates) {
        if (!rate.service || rate.service === '') {
            missingFieldsCount.service++;
            missingServices.push({ POL: rate.polName || '', POD: rate.podName || '' });
            if (rate.polName && rate.podName) {
                missingFieldsCount.missing_service_routes.push(`${rate.polName} -> ${rate.podName}`);
            }
        }
        if (!rate['Place of Delivery']) {
            missingFieldsCount.port_destination++;
        }
        if (!rate.POL) {
            missingFieldsCount.port_origin++;
            if (rate.polName) {
                missingFieldsCount.port_origin_names.push(rate.polName);
            }
        }
        if (!rate.POD) {
            missingFieldsCount.port_discharge++;
        }
    }

    // Step 4: Generate jsonRates, filtering out rates with empty POL or POD
    const jsonRates = [];
    for (const rate of finalRates) {
        if (!rate.POL || !rate.POD) {
            unprocessedRates.push({
                polName: rate.polName || '',
                podName: rate.podName || '',
                placeOfDelivery: rate['Place of Delivery'] || '',
                POL: rate.POL || '',
                POD: rate.POD || '',
                D20: rate.D20 || '',
                D40: rate.D40 || '',
                reason: !rate.POL ? 'Missing POL' : 'Missing POD'
            });
        } else {
            const jsonRate = {
                carrier: "653309",
                port_destination: rate['Place of Delivery'] || '',
                port_origin: rate.POL || '',
                port_discharge: rate.POD || '',
                port_origin_name: rate.polName || '',
                port_discharge_name: rate.podName || '',
                vf: contractEffectiveDate || '',
                vt: contractExpirationDate || '',
                transit_time: "",
                rate_source: "653309",
                "40p": rate.D40 || '',
                "40hqp": rate.D40 || '',
                "20p": rate.D20 || '',
                service: rate.service || '',
                service_name: rate.service_name || '',
                contract: "38",
                carrier_contract_number: "SVC3118"
            };
            jsonRates.push(jsonRate);
        }
    }

    // Step 5: Write UnprocessedRatesToEndpoint.csv (always, even if empty)
    const wsUnprocessed = XLSX.utils.json_to_sheet(unprocessedRates);
    const csvUnprocessed = XLSX.utils.sheet_to_csv(wsUnprocessed);
    const csvUnprocessedWithTs = `Timestamp,${timestamp}\n${csvUnprocessed}`;
    fs.writeFileSync('UnprocessedRatesToEndpoint.csv', csvUnprocessedWithTs, 'utf8');
    console.log(`UnprocessedRatesToEndpoint.csv generated/updated for ${sheetName} with ${unprocessedRates.length} rates.`);

    // Step 6: Write FinalRatesToEndpoint.json with timestamp, containing all jsonRates
    const finalRatesWithTimestamp = {
        timestamp: timestamp,
        rates: jsonRates
    };
    fs.writeFileSync('FinalRatesToEndpoint.json', JSON.stringify(finalRatesWithTimestamp, null, 2), 'utf8');
    console.log(`FinalRatesToEndpoint.json generated/updated for ${sheetName} with ${jsonRates.length} rates.`);

    // Step 6.1: Write SingleRateWithTimestamp.json with one rate and timestamp
    const singleRate = jsonRates.length > 0 ? jsonRates[0] : {};
    const singleRateWithTimestamp = {
        timestamp: timestamp,
        rate: singleRate
    };
    fs.writeFileSync('SingleRateWithTimestamp.json', JSON.stringify(singleRateWithTimestamp, null, 2), 'utf8');
    console.log(`SingleRateWithTimestamp.json generated with ${jsonRates.length > 0 ? '1 rate' : 'empty rate'}.`);

    // Step 7: Send up to 200 rates to endpoint
    let sentRateIndices = trackingState.sentRateIndices;
    const ratesSentThisRun = Math.min(jsonRates.length - sentRateIndices.length, maxRatesToSendPerRun);
    const unsentRates = jsonRates.filter((_, index) => !sentRateIndices.includes(index));
    const ratesToSend = unsentRates.slice(0, maxRatesToSendPerRun);
    const indicesToSend = jsonRates
        .map((rate, index) => ratesToSend.includes(rate) ? index : -1)
        .filter(index => index !== -1);

    if (ratesToSend.length > 0) {
        console.log(`Attempting to send ${ratesToSend.length} rates for ${sheetName}...`);
        const success = await sendRatesToEndpoint(ratesToSend, indicesToSend, sheetName);
        if (success) {
            sentRateIndices.push(...indicesToSend);
            trackingState.sentRateIndices = sentRateIndices;
            trackingState.totalRatesSent += ratesToSend.length;
            // Save trackingState to file
            fs.writeFileSync(trackingStateFile, JSON.stringify(trackingState, null, 2), 'utf8');
        }
    } else {
        console.log(`All rates have been sent to the endpoint for ${sheetName}.`);
    }

    // Calculate rates remaining
    let ratesRemainingToSend = jsonRates.length - sentRateIndices.length;
    let resetMessage = '';
    if (ratesRemainingToSend <= 0) {
        ratesRemainingToSend = 0;
        resetMessage = `All rates sent for ${sheetName}. Resetting sentRateIndices for the next sheet.`;
        console.log(resetMessage);
        sentRateIndices = [];
        trackingState.sentRateIndices = [];
        trackingState.totalRatesSent = 0; // Reset cumulative total for the sheet
        // Save reset trackingState
        fs.writeFileSync(trackingStateFile, JSON.stringify(trackingState, null, 2), 'utf8');
    }
    if (ratesRemainingToSend < 0) {
        console.warn(`Warning: Rates remaining to send is negative for ${sheetName}, indicating a tracking error. Resetting to 0.`);
        ratesRemainingToSend = 0;
        sentRateIndices = [];
        trackingState.sentRateIndices = [];
        // Save reset trackingState
        fs.writeFileSync(trackingStateFile, JSON.stringify(trackingState, null, 2), 'utf8');
    }

    // Step 8: Log statistics
    console.log(`\nRates Sent Statistics for ${sheetName}:`);
    console.log(`- Rates sent this run: ${ratesSentThisRun}`);
    console.log(`- Total rates sent: ${trackingState.totalRatesSent}`);
    console.log(`- Rates remaining to send: ${ratesRemainingToSend}`);
    console.log(`- Unprocessed rates (missing POL or POD): ${unprocessedRates.length}`);
    if (resetMessage) {
        console.log(`- ${resetMessage}`);
    }

    // Log missing port_origin_names
    if (missingFieldsCount.port_origin_names.length > 0) {
        console.log(`\nMissing port_origin_names for ${sheetName}:`, missingFieldsCount.port_origin_names);
    }

    // Log missing service ID routes
    if (missingFieldsCount.missing_service_routes.length > 0) {
        console.log(`\nMissing service ID routes for ${sheetName}:`, missingFieldsCount.missing_service_routes);
    }

    // Log skipped API rates
    if (skippedApiRates.length > 0) {
        console.log(`\nSkipped API Rates (invalid POL/POD) for ${sheetName}:`, JSON.stringify(skippedApiRates, null, 2));
    }

    // Statistics
    console.log(`\nService ID Fetch Statistics for ${sheetName}:`);
    console.log(`- Total rates received: ${finalRates.length}`);
    console.log(`- Rates processed: ${jsonRates.length}`);
    console.log(`- Rates with Service ID from Google Sheet: ${sheetMatchCount}`);
    console.log(`- Rates with Service ID from Betty Blocks API: ${apiMatchCount}`);
    console.log(`- Rates missing Service ID: ${missingFieldsCount.service}`);
    console.log(`- Rates missing port_destination: ${missingFieldsCount.port_destination}`);
    console.log(`- Rates missing port_origin: ${missingFieldsCount.port_origin}`);
    console.log(`- Rates missing port_discharge: ${missingFieldsCount.port_discharge}`);
    console.log(`- API calls made: ${apiCallCount}`);
    console.log(`- Service IDs received from API: ${apiMatchCount}`);
    console.log(`- Successful microservice updates (Google Sheet): ${microserviceUpdateCount}`);
    console.log(`- API calls failed (no service ID or error): ${apiFailedCount}`);
    console.log(`- API calls skipped (invalid POL/POD): ${apiSkippedCount}`);
}

module.exports = { createFinalRates };