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
        'BP PSW': [
            'YANTIAN', 'SHANGHAI', 'NINGBO', 'XIAMEN', 'KAOHSIUNG', 'QINGDAO', 'VUNG TAU',
            'TAIPEI', 'SINGAPORE', 'NANSHA', 'TIANJINXINGANG', 'PORT KELANG',
            'LAEM CHABANG', 'PUSAN', 'HAIPHONG'
        ],
        'PSW': ['LOS ANGELES', 'LONG BEACH', 'OAKLAND'],
        'LAX-LGB': ['LOS ANGELES', 'LONG BEACH'],
        'BP PNW': ['YANTIAN', 'HONG KONG', 'SHANGHAI', 'NINGBO', 'XIAMEN', 'KAOHSIUNG', 'PUSAN'],
        'PNW': ['SEATTLE', 'TACOMA'],
        'KHPNH': ['PHNOM PENH'],
    };
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
        output.push(rowObject);
    }

    let rates = output.map(rate => {
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
    const processedRates = finalRates.filter(rate => 
        rate.POL && rate.POD && (rate["Place of Delivery"] || !rate.originalPlaceOfDelivery)
    );

    console.log("Missing locations:", missingLocations);

    return {
        processedRates,
        preDictionaryRates,
        postDictionaryRates
    };
}

module.exports = { processUSWCSheet };