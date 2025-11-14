// coverSheet.js
const XLSX = require('xlsx');

function excelDateToJSDate(serial) {
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    return new Date(utcValue * 1000);
}

async function processCoverSheet(workbook) {
    const coverSheet = workbook.Sheets['Cover'];
    if (!coverSheet) {
        console.error("No 'Cover' sheet found in the workbook");
        process.exit(1);
    }

    const coverData = XLSX.utils.sheet_to_json(coverSheet, { header: 1 });
    console.log('First 5 rows of Cover sheet:', coverData.slice(0, 5));

    let contractEffectiveDate = null;
    let contractExpirationDate = null;

    for (const row of coverData) {
        if (!Array.isArray(row)) continue;
        
        const rowText = row.join(' ').toLowerCase();
        
        if (rowText.includes('contract effective date')) {
            const serial = row.find(cell => typeof cell === 'number');
            if (serial) {
                const date = excelDateToJSDate(serial);
                contractEffectiveDate = date.toISOString().split('T')[0];
            }
        }
        
        if (rowText.includes('contract expiration date')) {
            const serial = row.find(cell => typeof cell === 'number');
            if (serial) {
                const date = excelDateToJSDate(serial);
                contractExpirationDate = date.toISOString().split('T')[0];
            }
        }
    }

    return { contractEffectiveDate, contractExpirationDate };
}

module.exports = { processCoverSheet };