// serviceIDs.js
const XLSX = require('xlsx');

async function processServiceIDs(workbook) {
    try {
        const serviceSheetName = 'Sheet1';
        if (!workbook.SheetNames.includes(serviceSheetName)) {
            console.error("Sheet1 not found in the workbook");
            return;
        }

        const serviceWorksheet = workbook.Sheets[serviceSheetName];
        const serviceData = XLSX.utils.sheet_to_json(serviceWorksheet, { header: 1 });

        const serviceHeaderRow = serviceData.find(row => row && row.length > 0);
        if (!serviceHeaderRow) {
            console.error("No header row found in Sheet1");
            return;
        }

        const serviceColumnMap = {};
        const serviceColumns = ['Service ID', 'Service Name'];
        serviceColumns.forEach(col => {
            const idx = serviceHeaderRow.findIndex(header =>
                typeof header === 'string' && header.toLowerCase().includes(col.toLowerCase())
            );
            if (idx === -1) {
                console.warn(`Service column "${col}" not found in header row`);
            } else {
                serviceColumnMap[col] = idx;
            }
        });

        const serviceIDs = [];
        for (let i = 1; i < serviceData.length; i++) {
            const row = serviceData[i];
            if (!row || row.length === 0) continue;

            const serviceID = row[serviceColumnMap['Service ID']];
            const serviceName = row[serviceColumnMap['Service Name']];

            if (serviceID && serviceName) {
                serviceIDs.push({
                    'Service ID': serviceID,
                    'Service Name': serviceName
                });
            }
        }

        const newWorkbook = XLSX.utils.book_new();
        const newWorksheet = XLSX.utils.json_to_sheet(serviceIDs);
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'ServiceIDs');
        XLSX.writeFile(newWorkbook, 'serviceID.xlsx');
        console.log("serviceID.xlsx created with Service ID data");

        return serviceIDs;
    } catch (error) {
        console.error("Error processing Service IDs:", error);
        throw error;
    }
}

module.exports = { processServiceIDs };