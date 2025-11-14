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

// Fuzzy column finder
function findCol(label, headerRow) {
    let bestIdx;
    let minDistance = Infinity;
    headerRow.forEach((h, idx) => {
        if (typeof h === 'string') {
            const dist = levenshteinDistance(label.toLowerCase(), h.toLowerCase());
            if (dist < minDistance && dist <= 2) { // allow up to 2 edits
                minDistance = dist;
                bestIdx = idx;
            }
        }
    });
    return bestIdx;
}

function findLocationId(locationName, records) {
    if (!locationName) return { id: null, reason: 'Location name is empty or null' };
    const searchName = String(locationName).trim().toLowerCase();
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

// Extract 'General Surcharges (VALID)' descriptions from the ISC-US sheet
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
        console.warn('General Surcharges (VALID) key not found on ISC-US sheet.');
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
        console.warn('General Surcharges (VALID) header not found near key on ISC-US sheet.');
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

function toNumberOrNull(val) {
    if (val == null) return null;
    if (typeof val === 'number') return val;
    const s = String(val).trim();
    if (!s || /not\s*applicable/i.test(s)) return null;
    const cleaned = s.replace(/[,\s]/g, '');
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : null;
}

// ✅ Simplified tab detection
function detectISCUSTabName(workbook) {
    const names = workbook.SheetNames || [];
    const match = names.find(n => n.toUpperCase().includes("ISC-US"));
    return match || null;
}

async function processISCUSSheet(workbook, contractEffectiveDate, contractExpirationDate) {
    console.log('Processing ISC-US sheet...');

    if (!workbook || !workbook.SheetNames) {
        console.error('Invalid workbook provided');
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    const sheetName = detectISCUSTabName(workbook);
    if (!sheetName) {
        console.error("No sheet found containing 'ISC-US'");
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    console.log(`[ISC-US] Using sheet: ${sheetName}`);
    console.log(`[ISC-US] Available sheets: ${workbook.SheetNames.join(', ')}`);
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (!data || !Array.isArray(data)) {
        console.error('Invalid data in ISC-US sheet');
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    // Extract VALID general surcharges from this ISC-US sheet
    let generalSurchargesDescriptions = [];
    try {
        const validSurcharges = extractGeneralSurchargesValidFromRows(data);
        generalSurchargesDescriptions = validSurcharges.map(s => s.description).filter(Boolean);
        console.log('ISC-US General Surcharges (VALID) Descriptions:', JSON.stringify(generalSurchargesDescriptions));
    } catch (e) {
        console.warn('ISC-US General Surcharges extraction failed:', e && e.message ? e.message : e);
    }

    const requiredColumns = ['POL', 'POD', 'Place of Delivery', 'Curr', 'D20', 'D40'];

    const headerRowIndex = data.findIndex((row, idx) => {
        if (!row || row.length === 0) return false;
        const matchCount = row.filter(cell =>
            typeof cell === 'string' &&
            requiredColumns.some(col => cell.toLowerCase().includes(col.toLowerCase()))
        ).length;
        const lower = row.map(c => (typeof c === 'string' ? c.toLowerCase() : ''));
        const hasPOL = lower.some(c => c.includes('pol') || c.includes('port of loading'));
        const hasPOD = lower.some(c => c.includes('pod') || c.includes('port of discharge'));
        const ok = (hasPOL && hasPOD) || matchCount >= 2;
        if (ok) console.log(`[ISC-US] Header candidate at row ${idx + 1}:`, { hasPOL, hasPOD, matchCount });
        return ok;
    });

    if (headerRowIndex === -1) {
        console.error("No suitable header row found in ISC-US sheet");
        console.log("First 10 rows for debugging:\n", JSON.stringify(data.slice(0, 10), null, 2));
        return { processedRates: [], preDictionaryRates: [], postDictionaryRates: [] };
    }

    // Build composite header from this row and the next (to capture split headers)
    const headerRow1 = (data[headerRowIndex] || []).map(c => (typeof c === 'string' ? c.trim() : c));
    const headerRow2 = (data[headerRowIndex + 1] || []).map(c => (typeof c === 'string' ? c.trim() : c));
    const compositeLen = Math.max(headerRow1.length, headerRow2.length);
    const compositeHeader = Array.from({ length: compositeLen }, (_, i) => {
        const p1 = headerRow1[i] || '';
        const p2 = headerRow2[i] || '';
        return [p1, p2].filter(Boolean).join(' ').trim();
    });
    console.log(`[ISC-US] Header row index (0-based): ${headerRowIndex}`);
    console.log('[ISC-US] Composite header sample:', JSON.stringify(compositeHeader.slice(0, 25)));

    const columnMap = {};
    requiredColumns.forEach(col => {
        columnMap[col] = findCol(col, compositeHeader);
        if (columnMap[col] === undefined) {
            console.warn(`Column "${col}" not found in composite header`);
        }
    });
    columnMap['H40'] = findCol('H40', compositeHeader) ?? findCol('40HC', compositeHeader) ?? findCol('D40HC', compositeHeader);
    columnMap['H45'] = findCol('H45', compositeHeader) ?? findCol('45HC', compositeHeader) ?? findCol('D45HC', compositeHeader);
    console.log('[ISC-US] Column map:', JSON.stringify(columnMap));

    const dataStartRow = headerRowIndex + 2;
    console.log(`[ISC-US] Data scanning starts at row ${dataStartRow + 1} (0-based ${dataStartRow})`);

    const output = [];
    const routeIndices = new Map();
    const sampleExtracts = [];
    const pick = (row, idx) => (idx === undefined ? undefined : row[idx]);
    for (let i = dataStartRow; i < data.length; i++) {
        const row = data[i];
        if (!row || row.length === 0) continue;

        if (worksheet['!rows'] && worksheet['!rows'][i] && worksheet['!rows'][i].hidden) {
            continue;
        }

        const polVal = pick(row, columnMap['POL']);
        const podVal = pick(row, columnMap['POD']);
        const isPolEmpty = polVal == null || (typeof polVal === 'string' && polVal.trim() === "");
        const isPodEmpty = podVal == null || (typeof podVal === 'string' && podVal.trim() === "");
        if (isPolEmpty && isPodEmpty) {
            console.log(`[ISC-US] Hit empty POL+POD at row ${i + 1}. Stopping.`);
            break;
        }

        const rowObject = {};
        for (const [col, idx] of Object.entries(columnMap)) {
            if (idx !== undefined) rowObject[col] = row[idx] !== undefined ? row[idx] : null;
        }

        const routeKey = `${rowObject.POL || ''}->${rowObject.POD || ''}->${rowObject['Place of Delivery'] || ''}`;
        if (routeIndices.has(routeKey)) {
            routeIndices.get(routeKey).push(output.length);
        } else {
            routeIndices.set(routeKey, [output.length]);
        }

        output.push(rowObject);

        if (sampleExtracts.length < 8) {
            sampleExtracts.push({
                rowNumber: i + 1,
                POL: rowObject.POL,
                POD: rowObject.POD,
                Delivery: rowObject['Place of Delivery'],
                Curr: rowObject.Curr,
                D20: rowObject.D20,
                D40: rowObject.D40,
                H40: rowObject.H40,
                H45: rowObject.H45,
            });
        }
    }

    console.log(`[ISC-US] Extracted rows before dedupe: ${output.length}`);
    if (sampleExtracts.length) console.log('[ISC-US] Sample extracted rows:', JSON.stringify(sampleExtracts, null, 2));

    const keepIndices = new Set();
    routeIndices.forEach(indices => {
        const lastIndex = indices[indices.length - 1];
        keepIndices.add(lastIndex);
    });

    const filteredOutput = output.filter((_, i) => keepIndices.has(i));
    console.log(`[ISC-US] Rows after dedupe: ${filteredOutput.length}`);

    const excludedCities = [];

    let rates = filteredOutput.map(rate => {
        ['POL', 'POD', 'Place of Delivery'].forEach(key => {
            if (rate[key] && typeof rate[key] === 'string') {
                rate[key] = rate[key].split(',')[0].trim();
            }
        });
        if ('D20' in rate) rate.D20 = toNumberOrNull(rate.D20);
        if ('D40' in rate) rate.D40 = toNumberOrNull(rate.D40);
        if ('H40' in rate) rate.H40 = toNumberOrNull(rate.H40);
        if ('H45' in rate) rate.H45 = toNumberOrNull(rate.H45);
        return rate;
    });

    rates = rates.filter(row => {
        const pod = row['Place of Delivery'];
        if (!pod || typeof pod !== 'string') return true;
        return !excludedCities.some(city => pod.toLowerCase().trim() === String(city).toLowerCase().trim());
    });

    const preDictionaryRates = [...rates];
    const postDictionaryRates = [...rates];
    console.log(`[ISC-US] preDictionary: ${preDictionaryRates.length}, postDictionary: ${postDictionaryRates.length}`);

    const locationsFilePath = path.join(__dirname, 'LocationsBettyBlocks.json');
    if (!fs.existsSync(locationsFilePath)) {
        console.error("LocationsBettyBlocks.json does not exist. Please fetch it first.");
        return { processedRates: [], preDictionaryRates, postDictionaryRates };
    }
    const locationsData = JSON.parse(fs.readFileSync(locationsFilePath, 'utf8'));
    if (!locationsData.records || !Array.isArray(locationsData.records) || locationsData.records.length === 0) {
        console.error("LocationsBettyBlocks.json is empty or malformed.");
        return { processedRates: [], preDictionaryRates, postDictionaryRates };
    }

    const missingLocations = [];
    const unprocessedRates = [];
    let missingPolCount = 0, missingPodCount = 0, missingPlaceCount = 0;
    const finalRates = postDictionaryRates.map((rate, index) => {
        const polName = rate.POL;
        const podName = rate.POD;
        const placeOfDeliveryName = rate['Place of Delivery'] && typeof rate['Place of Delivery'] === 'string' && rate['Place of Delivery'].trim() !== '' ? rate['Place of Delivery'] : podName;

        const polResult = findLocationId(polName, locationsData.records);
        const polId = polResult.id;
        if (!polId) {
            missingPolCount++;
            missingLocations.push({ location: polName, type: 'POL', rateIndex: index, rate: `${polName} -> ${podName}`, reason: polResult.reason });
        }

        const podResult = findLocationId(podName, locationsData.records);
        const podId = podResult.id;
        if (!podId) {
            missingPodCount++;
            missingLocations.push({ location: podName, type: 'POD', rateIndex: index, rate: `${polName} -> ${podName}`, reason: podResult.reason });
        }

        let placeId = null;
        if (placeOfDeliveryName && typeof placeOfDeliveryName === 'string' && placeOfDeliveryName.trim() !== '') {
            const placeResult = findLocationId(placeOfDeliveryName, locationsData.records);
            placeId = placeResult.id;
            if (!placeId) {
                missingPlaceCount++;
                missingLocations.push({ location: placeOfDeliveryName, type: 'Place of Delivery', rateIndex: index, rate: `${polName} -> ${podName}`, reason: placeResult.reason });
            }
        } else {
            placeId = podId;
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
    console.log(`[ISC-US] processedRates: ${processedRates.length}. Missing POL=${missingPolCount}, POD=${missingPodCount}, Delivery=${missingPlaceCount}`);

    return { processedRates, preDictionaryRates, postDictionaryRates, uswcSurchargeDescriptions: generalSurchargesDescriptions };
}

module.exports = { processISCUSSheet };
