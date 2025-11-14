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
        'BP BAL': ['XIAMEN', 'HONG KONG', 'YANTIAN'],
        'SE ASIA BP GCFL': ['VUNG TAU', 'SINGAPORE'],
        'GULF': ['HOUSTON', 'MOBILE', 'NEW ORLEANS'],
        'BP JAPAN': ['NAGOYA', 'SHIMIZU', 'TOKYO', 'KOBE', 'YOKOHAMA', 'OSAKA', 'HIROSHIMA', 'MOJI', 'HAKATA/FUKUOKA'],
        'NCPRC BP EC': ['SHANGHAI', 'NINGBO', 'QINGDAO'],
        'SE ASIA BP EC': ['VUNG TAU', 'PORT KLANG', 'SINGAPORE', 'HAIPHONG'],
        'BP MIA': ['PORT KLANG', 'HAIPHONG', 'YANTIAN', 'NINGBO', 'SHANGHAI', 'XIAMEN'],
        'SPRC BP EC': ['YANTIAN', 'SHEKOU', 'XIAMEN', 'HONG KONG'],
        'BP GCFL': ['NINGBO', 'SHANGHAI', 'XIAMEN', 'YANTIAN', 'SHEKOU'],
        'BP BOS': ['SHANGHAI', 'NINGBO', 'QINGDAO'],
        'FAK GCFL': ['NINGBO', 'SHANGHAI', 'XIAMEN', 'YANTIAN', 'SHEKOU'],
        'BALTIMORE': ['BALTIMORE'],
        'NEW YORK': ['NEW YORK'],
        'NORFOLK': ['NORFOLK'],
        'SAVANNAH': ['SAVANNAH'],
        'CHARLESTON': ['CHARLESTON'],
        'MIAMI': ['MIAMI'],
        'TAMPA': ['TAMPA'],
        'HALIFAX': ['HALIFAX']
    };
}

// Extract 'General Surcharges (VALID)' descriptions from the USEC feeder sheet
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
        console.warn('General Surcharges (VALID) key not found on USEC Feeder sheet.');
        return [];
    }

    let headerRowIndex = -1;
    for (let i = keyRowIndex; i < Math.min(rows.length, keyRowIndex + 15); i++) {
        if (looksLikeHeader(rows[i])) { headerRowIndex = i; break; }
    }
    if (headerRowIndex === -1) {
        for (let i = keyRowIndex + 1; i < Math.min(rows.length, keyRowIndex + 10); i++) {
            const row = rows[i];
            if (Array.isArray(row) && row.some(cell => cell != null && String(cell).trim() !== '')) { headerRowIndex = i; break; }
        }
    }
    if (headerRowIndex === -1) {
        console.warn('General Surcharges (VALID) header not found near key on USEC Feeder sheet.');
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
    const searchName = locationName.trim().toLowerCase();
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
    const cityOnly = searchName.split(',')[0].trim();
    for (const record of records) {
        for (const key in record) {
            if (key === 'id') continue;
            const field = record[key];
            if (field && typeof field.value === 'string') {
                if (field.value.trim().toLowerCase() === cityOnly) {
                    return { id: record.id, reason: null };
                }
            }
        }
    }
    let closestMatch = null;
    let minDistance = Infinity;
    for (const record of records) {
        for (const key in record) {
            if (key === 'id') continue;
            const field = record[key];
            if (field && typeof field.value === 'string') {
                const distance = levenshteinDistance(searchName, field.value.trim().toLowerCase());
                if (distance < minDistance && distance <= 3) {
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

async function processUSECSheet(workbook, contractEffectiveDate, contractExpirationDate) {
    console.log('Processing USEC sheet...');

    if (!workbook || !workbook.SheetNames) {
        console.error('Invalid workbook provided');
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    const sheetName = workbook.SheetNames.find(name => name.includes('USEC'));
    if (!sheetName) {
        console.error("No sheet found containing 'USEC'");
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Extract VALID general surcharges from this USEC feeder sheet
    let generalSurchargesDescriptions = [];
    try {
        const validSurcharges = extractGeneralSurchargesValidFromRows(data);
        generalSurchargesDescriptions = validSurcharges.map(s => s.description).filter(Boolean);
        console.log('USEC Feeder General Surcharges (VALID) Descriptions:', JSON.stringify(generalSurchargesDescriptions));
    } catch (e) {
        console.warn('USEC Feeder General Surcharges extraction failed:', e && e.message ? e.message : e);
    }

    if (!data || !Array.isArray(data)) {
        console.error('Invalid data in USEC sheet');
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    const requiredColumns = ['POL', 'POD', 'Place of Delivery', 'Curr', 'D20', 'D40', 'Note'];
    const headerRowIndex = data.findIndex(row => row.some(cell => requiredColumns.includes(cell)));
    if (headerRowIndex === -1) {
        console.error("Could not find header row with required columns in USEC sheet");
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    const headerRow = data[headerRowIndex].map(cell => (typeof cell === 'string' ? cell.trim() : cell));
    const columnMap = {};
    requiredColumns.forEach(col => {
        const idx = headerRow.findIndex(header => typeof header === 'string' && header.includes(col));
        if (idx !== -1) columnMap[col] = idx;
    });

    const output = [];
    const routeIndices = new Map(); //  Added to track route (POL->POD->Place of Delivery) to list of indices
    for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        // Check for hidden rows
        if (worksheet['!rows'] && worksheet['!rows'][i] && worksheet['!rows'][i].hidden) {
            console.log(`Skipping hidden row ${i + 1}`);
            continue;
        }

        const polVal = row[columnMap['POL']];
        const podVal = row[columnMap['POD']];
        const isPolEmpty = polVal == null || (typeof polVal === 'string' && polVal.trim() === "");
        const isPodEmpty = podVal == null || (typeof podVal === 'string' && podVal.trim() === "");

        if (isPolEmpty && isPodEmpty) {
            console.log("Encountered a row with empty POL and POD; stopping processing.");
            break;
        }

        const rowObject = {};
        for (const [col, idx] of Object.entries(columnMap)) {
            rowObject[col] = row[idx] !== undefined ? row[idx] : null;
        }

        //  Define the route and track indices
        const routeKey = `${rowObject.POL || ''}->${rowObject.POD || ''}->${rowObject['Place of Delivery'] || ''}`;
        if (routeIndices.has(routeKey)) {
            routeIndices.get(routeKey).push(output.length);
        } else {
            routeIndices.set(routeKey, [output.length]);
        }

        output.push(rowObject);
    }

    //  Filter output to keep only the last occurrence of each route
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

    let rates = filteredOutput.map(rate => { // Use filteredOutput instead of output
        ['POL', 'POD'].forEach(key => {
            if (rate[key] && typeof rate[key] === 'string') {
                rate[key] = rate[key].split(',')[0].trim();
            }
        });
        return rate;
    });

    const preDictionaryRates = [...rates];

    const portDictionary = getPortGroupDictionary();
    rates = rates.map(rate => {
        const newRate = { ...rate };
        if (newRate.POL) {
            const portGroup = portDictionary[newRate.POL];
            if (portGroup) newRate.POL = portGroup;
        }
        if (newRate.POD) {
            const portGroup = portDictionary[newRate.POD];
            if (portGroup) newRate.POD = portGroup;
        }
        return newRate;
    });

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
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }
    const locationsData = JSON.parse(fs.readFileSync(locationsFilePath, 'utf8'));
    if (!locationsData.records || !Array.isArray(locationsData.records) || locationsData.records.length === 0) {
        console.error("LocationsBettyBlocks.json is empty or malformed.");
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    const missingLocations = [];
    const unprocessedRates = [];
    const finalRates = postDictionaryRates.map((rate, index) => {
        const polName = rate.POL;
        const podName = rate.POD;
        const placeOfDeliveryName = rate['Place of Delivery'] && typeof rate['Place of Delivery'] === 'string' && rate['Place of Delivery'].trim() !== '' ? rate['Place of Delivery'] : podName;

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
            placeId = podId;
            placeReason = placeId ? null : 'Place of Delivery empty and no POD ID available';
        }

        const newRate = {
            ...rate,
            polName,
            podName,
            originalPlaceOfDelivery: placeOfDeliveryName,
            'Place of Delivery': placeId,
            POL: polId,
            POD: podId
        };

        if (!polId || !podId || (!placeId && placeOfDeliveryName && placeOfDeliveryName.trim() !== '')) {
            unprocessedRates.push(newRate);
        }

        return newRate;
    });

    const processedRates = finalRates.filter(rate =>
        rate.POL && rate.POD && (rate['Place of Delivery'] || !rate.originalPlaceOfDelivery)
    );

    console.log('Missing locations:', missingLocations);

    // === NEW FEEDER LOGIC FOR USEC ===
    // Find the header row for the feeder table
    const feederRequiredColumns = ['COUNTRY', 'Place of Receipt', 'POL', 'D20', 'D40'];
    const feederHeaderRowIndex = data.findIndex(row => {
        if (!row || row.length === 0) return false;
        return feederRequiredColumns.every(col =>
            row.some(cell => typeof cell === 'string' && cell.replace(/\s+/g, '').toLowerCase().includes(col.replace(/\s+/g, '').toLowerCase()))
        );
    });

    let feederTable = [];
    if (feederHeaderRowIndex !== -1) {
        const feederHeaderRow = data[feederHeaderRowIndex].map(cell => (typeof cell === 'string' ? cell.trim() : cell));
        console.log('[DEBUG] Detected feeder header row:', feederHeaderRow);
        const feederColumnMap = {};
        feederRequiredColumns.forEach(col => {
            const idx = feederHeaderRow.findIndex(header =>
                typeof header === 'string' && header.replace(/\s+/g, '').toLowerCase().includes(col.replace(/\s+/g, '').toLowerCase())
            );
            if (idx !== -1) feederColumnMap[col] = idx;
        });
        console.log('[DEBUG] Feeder column map:', feederColumnMap);
        for (let i = feederHeaderRowIndex + 1; i < data.length; i++) {
            const row = data[i];
            if (!row || row.length === 0) continue;
            const polVal = row[feederColumnMap['POL']];
            const porVal = row[feederColumnMap['Place of Receipt']];
            const isPolEmpty = polVal == null || (typeof polVal === 'string' && polVal.trim() === "");
            const isPorEmpty = porVal == null || (typeof porVal === 'string' && porVal.trim() === "");
            if (isPolEmpty && isPorEmpty) break;
            const rowObject = {};
            Object.entries(feederColumnMap).forEach(([col, idx]) => {
                rowObject[col] = row[idx] !== undefined ? row[idx] : null;
            });
            feederTable.push(rowObject);
        }
        console.log(`[DEBUG] Extracted ${feederTable.length} feeder data rows.`);
        if (feederTable.length > 0) {
            console.log('[DEBUG] First feeder data row:', feederTable[0]);
        }
        if (feederTable.length > 1) {
            console.log('[DEBUG] Second feeder data row:', feederTable[1]);
        }
        // Normalize and deduplicate by Place of Receipt, keep lowest D20+D40
        const allFeederRows = feederTable.map(row => {
            const formattedPlaceOfReceipt = row['Place of Receipt'] && typeof row['Place of Receipt'] === 'string' ? row['Place of Receipt'].split(',')[0].trim() : row['Place of Receipt'];
            let placeOfReceiptId = null;
            if (formattedPlaceOfReceipt && locationsData && locationsData.records) {
                const result = findLocationId(formattedPlaceOfReceipt, locationsData.records);
                placeOfReceiptId = result.id;
            }
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
        const lowestByPlace = new Map();
        allFeederRows.forEach(row => {
            const key = row['Place of Receipt'];
            const grandTotal = row._d20 + row._d40;
            if (!lowestByPlace.has(key) || grandTotal < (lowestByPlace.get(key)._d20 + lowestByPlace.get(key)._d40)) {
                lowestByPlace.set(key, { D20: row.D20, D40: row.D40, _d20: row._d20, _d40: row._d40 });
            }
        });
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
        // === OVERWRITE Feeder.csv with only the latest USEC feeder rates ===
        const feederCsvPath = path.join(__dirname, 'Feeder.csv');
        const wsFeeder = XLSX.utils.json_to_sheet(feederRates);
        const csvFeeder = XLSX.utils.sheet_to_csv(wsFeeder);
        const timestamp = new Date().toISOString();
        const csvFeederWithTs = `Timestamp,${timestamp}\n${csvFeeder}`;
        fs.writeFileSync(feederCsvPath, csvFeederWithTs, 'utf8');
        console.log('Feeder.csv overwritten with latest USEC feeder rates.');
    }

    // Build merged processedRates: for each feeder row, find ALL base rates for matching POL, create merged rate per Place of Receipt
    const normalize = s => String(s || '').trim().toUpperCase();
    let feederRatesArr = [];
    try {
        const feederCsv = fs.readFileSync(path.join(__dirname, 'Feeder.csv'), 'utf8');
        const feederLines = feederCsv.split('\n').filter(Boolean);
        // Skip the timestamp line if present
        let startIdx = 0;
        if (feederLines[0].startsWith('Timestamp,')) startIdx = 2; // skip timestamp and header
        else startIdx = 1; // skip header only
        for (let i = startIdx; i < feederLines.length; i++) {
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
        return matchingBases.map(base => {
            let mergedD20 = (Number(base.D20) > 1 ? (Number(base.D20) || 0) + (Number(feeder.D20) || 0) : 0);
            let mergedD40 = (Number(base.D40) > 1 ? (Number(base.D40) || 0) + (Number(feeder.D40) || 0) : 0);
            if (!Number.isFinite(mergedD20)) mergedD20 = 0;
            if (!Number.isFinite(mergedD40)) mergedD40 = 0;
            return {
                ...base,
                POL: feeder['Place of Receipt ID'],
                D20: mergedD20,
                D40: mergedD40,
                feederD20: feeder.D20,
                feederD40: feeder.D40,
                baseD20: base.D20,
                baseD40: base.D40,
                feederPlaceOfReceipt: feeder['Place of Receipt'],
                feederPlaceOfReceiptID: feeder['Place of Receipt ID'],
                "Place of Receipt": feeder['Place of Receipt']
            };
        });
    });
    // Log the first merged rate
    if (mergedRates.length > 0) {
        console.log('\n=== FIRST MERGED PROCESSED RATE (Base + Feeder, USEC) ===');
        console.log(JSON.stringify(mergedRates[0], null, 2));
        console.log('=====================================');
    } else {
        console.log('\nNo merged processed rates found for USEC!');
    }
    // Write BaseandFeederRatesUSEC.csv
    try {
        const wsBaseFeeder = XLSX.utils.json_to_sheet(mergedRates);
        const csvBaseFeeder = XLSX.utils.sheet_to_csv(wsBaseFeeder);
        fs.writeFileSync('BaseandFeederRatesUSEC.csv', csvBaseFeeder, 'utf8');
        console.log('BaseandFeederRatesUSEC.csv created/updated.');
    } catch (e) {
        console.error('Could not write BaseandFeederRatesUSEC.csv:', e.message);
    }
    // Return mergedRates as processedRates
    return {
        processedRates: mergedRates,
        preDictionaryRates,
        postDictionaryRates,
        uswcSurchargeDescriptions: generalSurchargesDescriptions
    };
}

module.exports = { processUSECSheet };
