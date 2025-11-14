const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Simple Levenshtein distance for fuzzy matching
function levenshteinDistance(a, b) {
    const matrix = Array(b.length + 1).fill().map(() => Array(a.length + 1).fill(0));
    for (let i = 0; i <= a.length; i++) matrix[0][i] = i;
    for (let j = 0; j <= b.length; j++) matrix[j][0] = j;
    for (let j = 1; j <= b.length; j++) {
        for (let i = 1; i <= a.length; i++) {
            const indicator = a[i - 1] === b[j - 1] ? 0 : 1;
            matrix[j][i] = Math.min(
                matrix[j][i - 1] + 1,
                matrix[j - 1][i] + 1,
                matrix[j - 1][i - 1] + indicator
            );
        }
    }
    return matrix[b.length][a.length];
}

function getPortGroupDictionary() {
    return {
        'BP TW': ['KAOHSIUNG', 'TAIPEI'],
        'BP VAN': ['YANTIAN', 'XIAMEN', 'NINGBO', 'SHANGHAI', 'HONG KONG', 'SHEKOU'],
        'LAX-LGB': ['LOS ANGELES', 'LONG BEACH'],
        'BP SEA': ['SHANGHAI', 'YANTIAN', 'XIAMEN', 'NINGBO', 'SHEKOU', 'HONG KONG'],
        'BP TIW': ['YANTIAN', 'SHANGHAI', 'NINGBO'],
        'BP FUJI': ['TOKYO', 'NAGOYA', 'KOBE'],
        'BP LAX': ['NANSHA', 'YANTIAN', 'XIAMEN', 'SHANGHAI', 'NINGBO', 'QINGDAO', 'TIANJINXINGANG'],
        'BP OAK': ['YANTIAN', 'SHEKOU', 'SHANGHAI', 'NINGBO', 'QINGDAO'],
        'SE ASIA BP PSW': ['PORT KLANG', 'SINGAPORE', 'LAEM CHABANG', 'VUNG TAU', 'HAIPHONG'],
        'BP PRR': ['SHANGHAI', 'QINGDAO', 'TIANJINXINGANG'],
    };
}

// Extract 'General Surcharges (VALID)' descriptions from the USWC sheet (Feeder)
function extractGeneralSurchargesValidFromRows(rows) {
    const isKeyCell = (value) => {
        if (!value || typeof value !== 'string') return false;
        return value.replace(/\s+/g, '').toLowerCase().includes('generalsurchargesvalid');
    };
    const looksLikeHeader = (row) => {
        if (!Array.isArray(row)) return false;
        const normalized = row.map(c => typeof c === 'string' ? c.replace(/\s+/g, '').toLowerCase() : c);
        const hasCode = normalized.some(c => typeof c === 'string' && (c.includes('code') || c.includes('chargecode')));
        const hasDesc = normalized.some(c => typeof c === 'string' && (c.includes('description') || c.includes('charge')));
        const hasEff = normalized.some(c => typeof c === 'string' && (c.includes('effectivedate') || c.includes('startdate') || c.includes('effective')));
        const hasExp = normalized.some(c => typeof c === 'string' && (c.includes('expirationdate') || c.includes('enddate') || c.includes('expiration')));
        return (hasCode || hasDesc) && (hasEff || hasExp);
    };

    if (!rows || rows.length === 0) return [];
    const keyRowIndex = rows.findIndex(r => Array.isArray(r) && r.some(isKeyCell));
    if (keyRowIndex === -1) {
        console.warn('General Surcharges (VALID) key not found on USWC Feeder sheet.');
        return [];
    }

    let headerRowIndex = -1;
    for (let i = keyRowIndex; i < Math.min(rows.length, keyRowIndex + 15); i++) {
        if (looksLikeHeader(rows[i])) { headerRowIndex = i; break; }
    }
    if (headerRowIndex === -1) {
        for (let i = keyRowIndex + 1; i < Math.min(rows.length, keyRowIndex + 10); i++) {
            const row = rows[i];
            if (Array.isArray(row) && row.some(cell => cell != null && String(cell).trim() !== '')) {
                headerRowIndex = i;
                break;
            }
        }
    }
    if (headerRowIndex === -1) {
        console.warn('General Surcharges (VALID) header not found near key on USWC Feeder sheet.');
        return [];
    }

    const headerRow = rows[headerRowIndex].map(c => typeof c === 'string' ? c.trim() : c);
    const lowerHeader = headerRow.map(h => typeof h === 'string' ? h.toLowerCase() : h);
    const idxEffective = lowerHeader.findIndex(h => typeof h === 'string' && (h.includes('effective') || h.includes('start')));
    const idxExpiration = lowerHeader.findIndex(h => typeof h === 'string' && (h.includes('expiration') || h.includes('end')));
    const idxApplicable = lowerHeader.findIndex(h => typeof h === 'string' && (h.includes('applicable') || h.includes('applicability') || h.includes('status')));

    const valid = [];
    let sawAnyRow = false;
    for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!Array.isArray(row)) continue;
        const joined = row.filter(Boolean).map(v => String(v)).join(' ').toLowerCase();
        if (joined.includes('general surcharges not valid') || joined.includes('generalsurchargesnotvalid')) break;
        if (row.every(cell => cell == null || String(cell).trim() === '')) {
            if (sawAnyRow) break;
            continue;
        }
        sawAnyRow = true;

        let applicable = false;
        let notApplicable = false;
        if (idxApplicable !== -1) {
            const cell = row[idxApplicable];
            const cellVal = typeof cell === 'string' ? cell.trim().toLowerCase() : String(cell || '').trim().toLowerCase();
            applicable = ['applicable', 'yes', 'y'].some(t => cellVal === t || cellVal.startsWith('applicable'));
            notApplicable = cellVal.includes('not applicable') || ['no', 'n', 'notapplicable', 'na', 'n/a'].includes(cellVal);
        } else {
            applicable = row.some(cell => typeof cell === 'string' && cell.trim().toLowerCase() === 'applicable');
            notApplicable = row.some(cell => typeof cell === 'string' && cell.toLowerCase().includes('not applicable'));
        }
        if (!applicable || notApplicable) continue;

        const firstTwo = row.filter(c => c != null && String(c).trim() !== '').slice(0, 2).map(String);
        const code = firstTwo[0] ? firstTwo[0].trim() : '';
        const description = firstTwo[1] ? firstTwo[1].trim() : '';
        const effectiveDate = idxEffective !== -1 && row[idxEffective] != null ? String(row[idxEffective]).trim() : '';
        const expirationDate = idxExpiration !== -1 && row[idxExpiration] != null ? String(row[idxExpiration]).trim() : '';
        if (code || description) valid.push({ code, description, effectiveDate, expirationDate });
    }

    return valid;
}

function findLocationId(locationName, records) {
    if (!locationName) return { id: null, reason: 'Location name is empty or null' };
    // Alias mapping for common alternate names
    const aliasMap = {
        'HO CHI MINH CITY': 'Ho Chi Minh',
        // Add more aliases here as needed
    };
    let searchName = locationName.trim();
    if (aliasMap[searchName.toUpperCase()]) {
        searchName = aliasMap[searchName.toUpperCase()];
    }
    searchName = searchName.toLowerCase();
    for (const record of records) {
        for (const key in record) {
            if (key === 'id') continue;
            const field = record[key];
            if (field && typeof field.value === 'string') {
                if (field.value.trim().toLowerCase() === searchName) {
                    return { id: record.id, reason: null };
                }
            }
        }
    }
    // Find close matches for suggestions
    let closestMatch = null;
    let minDistance = Infinity;
    for (const record of records) {
        for (const key in record) {
            if (key === 'id') continue;
            const field = record[key];
            if (field && typeof field.value === 'string') {
                const distance = levenshteinDistance(searchName, field.value.trim().toLowerCase());
                if (distance < minDistance && distance <= 3) { // Threshold for close match
                    minDistance = distance;
                    closestMatch = field.value;
                }
            }
        }
    }
    const reason = closestMatch
        ? `No exact match found for "${locationName}". Closest match: "${closestMatch}"`
        : `No match found for "${locationName}" in LocationsBettyBlocks.json`;
    return { id: null, reason };
}

async function processUSWCSheet(workbook, contractEffectiveDate, contractExpirationDate) {
    const sheetName = workbook.SheetNames.find(name => name.includes('USWC'));    
    if (!sheetName) {
        console.error("No sheet found containing 'USWC'");
        process.exit(1);
    }

    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Extract VALID general surcharges from this USWC feeder sheet
    let generalSurchargesDescriptions = [];
    try {
        const validSurcharges = extractGeneralSurchargesValidFromRows(data);
        generalSurchargesDescriptions = validSurcharges.map(s => s.description).filter(Boolean);
        console.log('USWC Feeder General Surcharges (VALID) Descriptions:', JSON.stringify(generalSurchargesDescriptions));
    } catch (e) {
        console.warn('USWC Feeder General Surcharges extraction failed:', e && e.message ? e.message : e);
    }

    const requiredColumns = ['POL', 'POD', 'Place of Delivery', 'Curr', 'D20', 'D40'];
    const headerRowIndex = data.findIndex(row => {
        if (!row || row.length === 0) return false;
        const containsFakBullets = row.some(cell =>
            typeof cell === 'string' &&
            cell.replace(/\s+/g, '').toUpperCase().includes('FAK/BULLETS'.replace(/\s+/g, '').toUpperCase())
        );
        const containsOtherHeader = row.some(cell =>
            typeof cell === 'string' && requiredColumns.some(col => cell.toLowerCase().includes(col.toLowerCase()))
        );
        return containsFakBullets && containsOtherHeader;
    });

    if (headerRowIndex === -1) {
        process.exit(1);
    }

    const headerRow = data[headerRowIndex].map(cell => (typeof cell === 'string' ? cell.trim() : cell));

    const columnMap = {};
    requiredColumns.forEach(col => {
        const idx = headerRow.findIndex(header =>
            typeof header === 'string' && header.toLowerCase().includes(col.toLowerCase())
        );
        if (idx === -1) {
            console.warn(`Column "${col}" not found in header row`);
        } else {
            columnMap[col] = idx;
        }
    });

    const output = [];
    const routeIndices = new Map(); // Map to track route (POL->POD->Place of Delivery) to list of indices

    for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;
        
        const polVal = row[columnMap['POL']];
        const podVal = row[columnMap['POD']];
        const isPolEmpty = polVal == null || (typeof polVal === 'string' && polVal.trim() === "");
        const isPodEmpty = podVal == null || (typeof polVal === 'string' && polVal.trim() === "");
        
        if (isPolEmpty && isPodEmpty) {
            console.log("Encountered a row with empty POL and POD; stopping processing.");
            break;
        }
        
        const rowObject = {};
        for (const [col, idx] of Object.entries(columnMap)) {
            rowObject[col] = row[idx] !== undefined ? row[idx] : null;
        }

        // Define the route as POL->POD->Place of Delivery (removing old price)
        const routeKey = `${rowObject.POL || ''}->${rowObject.POD || ''}->${rowObject['Place of Delivery'] || ''}`;
        if (routeIndices.has(routeKey)) {
            routeIndices.get(routeKey).push(output.length);
        } else {
            routeIndices.set(routeKey, [output.length]);
        }

        output.push(rowObject);
    }

    console.log(`Total rows in output: ${output.length}`);
   

    // Filter output to keep only the last occurrence of each route
    const filteredOutput = [];
    const keepIndices = new Set();
    routeIndices.forEach((indices, routeKey) => {
        const lastIndex = indices[indices.length - 1]; // Keep the last occurrence (newest price)
        keepIndices.add(lastIndex);
    });

    for (let i = 0; i < output.length; i++) {
        if (keepIndices.has(i)) {
            filteredOutput.push(output[i]);
        }
    }

    console.log(`Total rows in filteredOutput: ${filteredOutput.length}`);
   
    // Filter out rows where Place of Delivery is Baltimore, Charlotte, New York, or Norfolk
    const citiesToIgnore = ['baltimore', 'charlotte', 'new york', 'norfolk'];
    const filteredByPlace = filteredOutput.filter(row => {
        const placeOfDelivery = row['Place of Delivery'];
        if (!placeOfDelivery || typeof placeOfDelivery !== 'string') return true; // Keep rows with no Place of Delivery
        return !citiesToIgnore.some(city => 
            placeOfDelivery.toLowerCase().trim() === city.toLowerCase()
        );
    });

    

    let rates = filteredByPlace.map(rate => {
        ['POL', 'POD', 'Place of Delivery'].forEach(key => {
            if (rate[key] && typeof rate[key] === 'string') {
                rate[key] = rate[key].split(',')[0].trim();
            }
        });
        return rate;
    });

  

    // Capture rates before dictionary replacement
    const preDictionaryRates = [...rates];

    

    const portDictionary = getPortGroupDictionary();
    rates = rates.map(rate => {
        const newRate = { ...rate };
        if (newRate.POL) {
            const portGroup = portDictionary[newRate.POL];
            if (portGroup) {
                newRate.POL = portGroup;
            }
        }
        if (newRate.POD) {
            const portGroup = portDictionary[newRate.POD];
            if (portGroup) {
                newRate.POD = portGroup;
            }
        }
        return newRate;
    });

    // Capture rates after dictionary replacement
    const postDictionaryRates = [];
    rates.forEach(rate => {
        const polValues = Array.isArray(rate.POL) ? rate.POL : [rate.POL];
        const podValues = Array.isArray(rate.POD) ? rate.POD : [rate.POD];
        polValues.forEach(pol => {
            podValues.forEach(pod => {
                const newRateRecord = { ...rate, POL: pol, POD: pod };
                postDictionaryRates.push(newRateRecord);
            });
        });
    });

   

    const locationsFilePath = path.join(__dirname, 'LocationsBettyBlocks.json');
    if (!fs.existsSync(locationsFilePath)) {
        console.error("LocationsBettyBlocks.json does not exist. Please fetch it first.");
        process.exit(1);
    }
    const locationsData = JSON.parse(fs.readFileSync(locationsFilePath, 'utf8'));
    if (!locationsData.records || !Array.isArray(locationsData.records) || locationsData.records.length === 0) {
        console.error("LocationsBettyBlocks.json is empty or malformed.");
        process.exit(1);
    }

    const missingLocations = [];
    const unprocessedRates = [];
    const finalRates = postDictionaryRates.map((rate, index) => {
        const polName = rate.POL;
        const podName = rate.POD;
        const placeOfDeliveryName = rate["Place of Delivery"];

        const polResult = findLocationId(polName, locationsData.records);
        const polId = polResult.id;
        if (!polId) {
            missingLocations.push({ 
                location: polName, 
                type: 'POL', 
                rateIndex: index, 
                rate: `${polName} -> ${podName}`, 
                reason: polResult.reason 
            });
        }

        const podResult = findLocationId(podName, locationsData.records);
        const podId = podResult.id;
        if (!podId) {
            missingLocations.push({ 
                location: podName, 
                type: 'POD', 
                rateIndex: index, 
                rate: `${polName} -> ${podName}`, 
                reason: podResult.reason 
            });
        }

        let placeId = null;
        let placeReason = null;
        if (placeOfDeliveryName && typeof placeOfDeliveryName === 'string' && placeOfDeliveryName.trim() !== '') {
            const placeResult = findLocationId(placeOfDeliveryName, locationsData.records);
            placeId = placeResult.id;
            placeReason = placeResult.reason;
            if (!placeId) {
                missingLocations.push({ 
                    location: placeOfDeliveryName, 
                    type: 'Place of Delivery', 
                    rateIndex: index, 
                    rate: `${polName} -> ${podName}`, 
                    reason: placeResult.reason 
                });
            }
        } else {
            placeId = podId; // Use POD ID only if Place of Delivery is empty
            placeReason = placeId ? null : 'Place of Delivery empty and no POD ID available';
        }

        const newRate = {
            ...rate,
            polName,
            podName,
            originalPlaceOfDelivery: placeOfDeliveryName, // Keep original for reference
            POL: polId,
            POD: podId,
            "Place of Delivery": placeId
        };

        // If any ID is missing, add to unprocessedRates
        if (!polId || !podId || (!placeId && placeOfDeliveryName && placeOfDeliveryName.trim() !== '')) {
            unprocessedRates.push(newRate);
        }

        return newRate;
    });

    // Filter out unprocessed rates from finalRates
    console.log("\nChecking for rates with unwanted cities...");
    const processedRates = finalRates.filter(rate => {
        // Log if we find a rate with unwanted cities
        const unwantedPortIds = ["649220", "657528", "656284", "657301"];
        if (unwantedPortIds.includes(String(rate["Place of Delivery"]))) {
        }
        return rate.POL && rate.POD && (rate["Place of Delivery"] || !rate.originalPlaceOfDelivery);
    }).filter(rate => {
        const unwantedPortIds = ["649220", "657528", "656284", "657301", "660208"]; // blocking new york,st louis, charlotte,norfolk   baltimore, charlotte, new york, norfolk and st louis
        return !unwantedPortIds.includes(String(rate["Place of Delivery"]));
    });

    // Log the first processed rate to show the structure
    if (processedRates.length > 0) {
        console.log("\n=== FIRST PROCESSED RATE STRUCTURE ===");
        console.log(JSON.stringify(processedRates[0], null, 2));
        console.log("=====================================\n");
    } else {
        console.log("\nNo processed rates found!");
    }

    // === NEW TABLE EXTRACTION ===
    // Find the header row for the new table
    const newTableRequiredColumns = ['COUNTRY', 'Place of Receipt', 'POL', 'D20', 'D40'];
    const newTableHeaderRowIndex = data.findIndex(row => {
        if (!row || row.length === 0) return false;
        // Only require COUNTRY, Place of Receipt, POL to find the header
        return ['COUNTRY', 'Place of Receipt', 'POL'].every(col =>
            row.some(cell => typeof cell === 'string' && cell.replace(/\s+/g, '').toLowerCase().includes(col.replace(/\s+/g, '').toLowerCase()))
        );
    });

    let newTable = [];
    if (newTableHeaderRowIndex !== -1) {
        const newTableHeaderRow = data[newTableHeaderRowIndex].map(cell => (typeof cell === 'string' ? cell.trim() : cell));
        // Map columns (look for all 4 columns)
        const newTableColumnMap = {};
        ['POL', 'Place of Receipt', 'D20', 'D40'].forEach(col => {
            const idx = newTableHeaderRow.findIndex(header =>
                typeof header === 'string' && header.replace(/\s+/g, '').toLowerCase().includes(col.replace(/\s+/g, '').toLowerCase())
            );
            if (idx !== -1) newTableColumnMap[col] = idx;
        });
        // Collect rows
        for (let i = newTableHeaderRowIndex + 1; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;
            const polVal = row[newTableColumnMap['POL']];
            const porVal = row[newTableColumnMap['Place of Receipt']];
            const isPolEmpty = polVal == null || (typeof polVal === 'string' && polVal.trim() === "");
            const isPorEmpty = porVal == null || (typeof porVal === 'string' && porVal.trim() === "");
            if (isPolEmpty && isPorEmpty) break;
            // Build row object
            const rowObject = {};
            Object.entries(newTableColumnMap).forEach(([col, idx]) => {
                rowObject[col] = row[idx] !== undefined ? row[idx] : null;
            });
            newTable.push(rowObject);
        }
        // Log the new table
        // console.log("\n=== NEW TABLE (COUNTRY, Place of Receipt, POL, D20, D40) ===");
        if (newTable.length > 0) {
            // console.log(JSON.stringify(newTable, null, 2));
        } else {
            console.log("No rows found in new table.");
        }
        console.log("===============================================\n");

        // === For each Place of Receipt, find the lowest D20+D40, and use those values for all rows with that Place of Receipt ===
        // First, build all rows with formatted values and ID
        const allFeederRows = newTable.map(row => {
            const formattedPlaceOfReceipt = row['Place of Receipt'] && typeof row['Place of Receipt'] === 'string' ? row['Place of Receipt'].split(',')[0].trim() : row['Place of Receipt'];
            let placeOfReceiptId = null;
            if (formattedPlaceOfReceipt && locationsData && locationsData.records) {
                const result = findLocationId(formattedPlaceOfReceipt, locationsData.records);
                placeOfReceiptId = result.id;
            }
            // Parse D20 and D40 as numbers, treat as 0 if missing or not a number
            const d20 = row.D20 && !isNaN(Number(row.D20)) ? Number(row.D20) : 0;
            const d40 = row.D40 && !isNaN(Number(row.D40)) ? Number(row.D40) : 0;
            return {
                POL: row.POL && typeof row.POL === 'string' ? row.POL.split(',')[0].trim() : row.POL,
                'Place of Receipt': formattedPlaceOfReceipt,
                D20: row.D20,
                D40: row.D40,
                'Place of Receipt ID': placeOfReceiptId,
                _d20: d20,
                _d40: d40
            };
        });
        // Find the lowest D20+D40 for each Place of Receipt
        const lowestByPlace = new Map();
        allFeederRows.forEach(row => {
            const key = row['Place of Receipt'];
            const grandTotal = row._d20 + row._d40;
            if (!lowestByPlace.has(key) || grandTotal < (lowestByPlace.get(key)._d20 + lowestByPlace.get(key)._d40)) {
                lowestByPlace.set(key, { D20: row.D20, D40: row.D40, _d20: row._d20, _d40: row._d40 });
            }
        });
        // For each row, set D20/D40 to the lowest for its Place of Receipt
        let feederRates = allFeederRows.map(row => {
            const lowest = lowestByPlace.get(row['Place of Receipt']);
            return {
                POL: row.POL,
                'Place of Receipt': row['Place of Receipt'],
                D20: lowest ? lowest.D20 : row.D20,
                D40: lowest ? lowest.D40 : row.D40,
                'Place of Receipt ID': row['Place of Receipt ID']
            };
        });
        // Deduplicate: keep only the first occurrence for each unique 'Place of Receipt'
        const seenPlaces = new Set();
        feederRates = feederRates.filter(row => {
            if (seenPlaces.has(row['Place of Receipt'])) return false;
            seenPlaces.add(row['Place of Receipt']);
            return true;
        });
        if (feederRates.length > 0) {
            const XLSX = require('xlsx');
            const fs = require('fs');
            const wsFeeder = XLSX.utils.json_to_sheet(feederRates);
            const csvFeeder = XLSX.utils.sheet_to_csv(wsFeeder);
            // Add timestamp as the first line
            const timestamp = new Date().toISOString();
            const csvFeederWithTs = `Timestamp,${timestamp}\n${csvFeeder}`;
            fs.writeFileSync('Feeder.csv', csvFeederWithTs, 'utf8');
            console.log('Feeder.csv created/updated with new table data.');
        }
    } else {
        console.log("\nNo header row found for new table (COUNTRY, Place of Receipt, POL).\n");
    }

    // === NEW processedRates with Feeder charges ===
    // Build merged processedRates: for each feeder row, find ALL base rates for matching POL, create merged rate per Place of Receipt
    const normalize = s => String(s || '').trim().toUpperCase();
    let feederRatesArr = [];
    try {
        const feederCsv = fs.readFileSync(path.join(__dirname, 'Feeder.csv'), 'utf8');
        const feederLines = feederCsv.split('\n').filter(Boolean);
        const feederHeader = feederLines[0].split(',');
        for (let i = 1; i < feederLines.length; i++) {
            const cols = feederLines[i].split(',');
            feederRatesArr.push({
                POL: cols[0],
                'Place of Receipt': cols[1],
                D20: Number(cols[2]) || 0,
                D40: Number(cols[3]) || 0,
                'Place of Receipt ID': cols[4] || null
            });
        }
    } catch (e) {
        console.error('Could not load Feeder.csv:', e.message);
    }
    // For each feeder row, merge with ALL matching base rates
    const mergedRates = feederRatesArr.flatMap(feeder => {
        if (!feeder['Place of Receipt ID']) return [];
        const matchingBases = processedRates.filter(rate => normalize(rate.polName) === normalize(feeder.POL));
        return matchingBases.map(base => ({
            ...base,
            POL: feeder['Place of Receipt ID'],
            D20: (Number(base.D20) || 0) + (Number(feeder.D20) || 0),
            D40: (Number(base.D40) || 0) + (Number(feeder.D40) || 0),
            feederD20: feeder.D20,
            feederD40: feeder.D40,
            baseD20: base.D20,
            baseD40: base.D40,
            feederPlaceOfReceipt: feeder['Place of Receipt'],
            feederPlaceOfReceiptID: feeder['Place of Receipt ID'],
            "Place of Receipt": feeder['Place of Receipt']
        }));
    });

    // Log the first merged rate
    if (mergedRates.length > 0) {
        console.log('\n=== FIRST MERGED PROCESSED RATE (Base + Feeder) ===');
        console.log(JSON.stringify(mergedRates[0], null, 2));
        console.log('=====================================');
    } else {
        console.log('\nNo merged processed rates found!');
    }

    // Write BaseandFeederRates sheet
    try {
        const wsBaseFeeder = XLSX.utils.json_to_sheet(mergedRates);
        const csvBaseFeeder = XLSX.utils.sheet_to_csv(wsBaseFeeder);
        fs.writeFileSync('BaseandFeederRates.csv', csvBaseFeeder, 'utf8');
        console.log('BaseandFeederRates.csv created/updated.');
    } catch (e) {
        console.error('Could not write BaseandFeederRates.csv:', e.message);
    }

    // Return mergedRates as processedRates
    return {
        processedRates: mergedRates,
        preDictionaryRates,
        postDictionaryRates,
        uswcSurchargeDescriptions: generalSurchargesDescriptions
    };
}

module.exports = { processUSWCSheet };
