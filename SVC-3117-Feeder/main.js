const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');
const { login } = require('./auth');
const { fetchAndSaveLocations } = require('./locations');
const { processCoverSheet } = require('./coverSheet');
const { processUSWCSheet } = require('./uswcSheet');
const { processServiceIDs } = require('./serviceIDs');
const { createFinalRates } = require('./finalRates');
const { processUSECSheet } = require('./usecSheet');

async function downloadRatesheet() {
    const url = 'https://www.primefreight.com/cma_rates/ratesheet.xlsx';
    const localPath = './ratesheet.xlsx';
    
    try {
        console.log('Downloading ratesheet from:', url);
        const response = await axios({
            method: 'GET',
            url: url,
            responseType: 'arraybuffer'
        });
        
        fs.writeFileSync(localPath, response.data);
        console.log('Ratesheet downloaded successfully to:', localPath);
        return localPath;
    } catch (error) {
        console.error('Error downloading ratesheet:', error.message);
        throw error;
    }
}

async function main() {
    const processMode = 'USWC'; // Can be 'USWC', 'USEC', or 'BOTH'
    const timestamp = new Date().toISOString();

    try {
        console.log(`Starting rate sheet processing for ${processMode}...`);

        // Authentication
        await login();

        // Fetch and save locations
        await fetchAndSaveLocations();

        // Download and read the workbook
        const ratesheetPath = await downloadRatesheet();
        const workbook = XLSX.readFile(ratesheetPath);

        // Process cover sheet for contract dates
        const { contractEffectiveDate, contractExpirationDate } = await processCoverSheet(workbook);
        let uswcRates = [];
        let usecRates = [];
        let uswcPreDictRates = [];
        let usecPreDictRates = [];
        let uswcPostDictRates = [];
        let usecPostDictRates = [];
        let uswcSurchargeDescriptions = [];

        // Process USWC rates
        if (processMode === 'USWC' || processMode === 'BOTH') {
            const uswcResult = await processUSWCSheet(workbook, contractEffectiveDate, contractExpirationDate);
            uswcRates = uswcResult.processedRates;
            uswcPreDictRates = uswcResult.preDictionaryRates;
            uswcPostDictRates = uswcResult.postDictionaryRates;
            uswcSurchargeDescriptions = uswcResult.uswcSurchargeDescriptions || [];

            // Write PreDictionaryRates.csv for USWC
            const wsUSWCPre = XLSX.utils.json_to_sheet(uswcPreDictRates);
            const csvUSWCPre = XLSX.utils.sheet_to_csv(wsUSWCPre);
            const csvUSWCPreWithTs = `Timestamp,${timestamp}\n${csvUSWCPre}`;
            fs.writeFileSync('PreDictionaryRates.csv', csvUSWCPreWithTs, 'utf8');

            // Write PostDictionaryRates.csv for USWC
            const wsUSWCPost = XLSX.utils.json_to_sheet(uswcPostDictRates);
            const csvUSWCPost = XLSX.utils.sheet_to_csv(wsUSWCPost);
            const csvUSWCPostWithTs = `Timestamp,${timestamp}\n${csvUSWCPost}`;
            fs.writeFileSync('PostDictionaryRates.csv', csvUSWCPostWithTs, 'utf8');
        }

        // Process USEC rates
        if (processMode === 'USEC' || processMode === 'BOTH') {
            const usecResult = await processUSECSheet(workbook, contractEffectiveDate, contractExpirationDate);
            usecRates = usecResult.processedRates;
            usecPreDictRates = usecResult.preDictionaryRates;
            usecPostDictRates = usecResult.postDictionaryRates;
            // Merge surcharge descriptions from USEC feeder
            const usecSurcharges = usecResult.uswcSurchargeDescriptions || [];
            if (usecSurcharges.length > 0) {
                const merged = Array.from(new Set([...(uswcSurchargeDescriptions || []), ...usecSurcharges]));
                uswcSurchargeDescriptions = merged;
            }

            // Write PreDictionaryRates.csv for USEC (append or create)
            const wsUSECPre = XLSX.utils.json_to_sheet(usecPreDictRates);
            const csvUSECPre = XLSX.utils.sheet_to_csv(wsUSECPre);
            const csvUSECPreWithTs = `Timestamp,${timestamp}\n${csvUSECPre}`;
            const preDictFile = 'PreDictionaryRates.csv';
            if (fs.existsSync(preDictFile) && processMode === 'BOTH') {
                const existingPreDict = fs.readFileSync(preDictFile, 'utf8').split('\n').slice(1).join('\n');
                fs.writeFileSync(preDictFile, `Timestamp,${timestamp}\n${existingPreDict}\n${csvUSECPre.split('\n').slice(1).join('\n')}`, 'utf8');
            } else {
                fs.writeFileSync(preDictFile, csvUSECPreWithTs, 'utf8');
            }

            // Write PostDictionaryRates.csv for USEC (append or create)
            const wsUSECPost = XLSX.utils.json_to_sheet(usecPostDictRates);
            const csvUSECPost = XLSX.utils.sheet_to_csv(wsUSECPost);
            const csvUSECPostWithTs = `Timestamp,${timestamp}\n${csvUSECPost}`;
            const postDictFile = 'PostDictionaryRates.csv';
            if (fs.existsSync(postDictFile) && processMode === 'BOTH') {
                const existingPostDict = fs.readFileSync(postDictFile, 'utf8').split('\n').slice(1).join('\n');
                fs.writeFileSync(postDictFile, `Timestamp,${timestamp}\n${existingPostDict}\n${csvUSECPost.split('\n').slice(1).join('\n')}`, 'utf8');
            } else {
                fs.writeFileSync(postDictFile, csvUSECPostWithTs, 'utf8');
            }
        }

        // Combine rates
        const allRates = [...uswcRates, ...usecRates];

        // Process Service IDs
        await processServiceIDs(workbook);

        // Generate final rates for endpoint
        await createFinalRates(allRates, contractEffectiveDate, contractExpirationDate, processMode, uswcSurchargeDescriptions);

        console.log(`Processing completed successfully for ${processMode}.`);

    } catch (error) {
        console.error(`Error in main process for ${processMode}:`, error);
        process.exit(1);
    }
}

main();
