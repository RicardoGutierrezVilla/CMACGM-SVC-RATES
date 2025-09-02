const XLSX = require('xlsx');
const axios = require('axios');
const { login } = require('./auth');
const { fetchAndSaveLocations } = require('./locations');
const { processCoverSheet } = require('./coverSheet');
const { processUSWCSheet } = require('./uswcSheet');
const { processUSECSheet } = require('./usecSheet');
const { processServiceIDs } = require('./serviceIDs');
const { createFinalRates } = require('./finalRates');

async function fetchRatesheetFromURL() {
    const url = 'https://www.primefreight.com/cma_rates/ratesheet.xlsx';
    console.log(`Fetching ratesheet from: ${url}`);
    
    try {
        const response = await axios.get(url, {
            responseType: 'arraybuffer',
            timeout: 30000 // 30 second timeout
        });
        
        console.log('Ratesheet fetched successfully');
        return XLSX.read(response.data, { type: 'buffer' });
    } catch (error) {
        console.error('Error fetching ratesheet from URL:', error.message);
        throw new Error(`Failed to fetch ratesheet from ${url}: ${error.message}`);
    }
}

async function main() {
    const processMode = 'BOTH'; // Can be 'USWC', 'USEC', or 'BOTH'
    const timestamp = new Date().toISOString();

    try {
        console.log(`Starting rate sheet processing for ${processMode}...`);

        // Authentication
        await login();

        // Fetch and save locations
        await fetchAndSaveLocations();

        // Fetch the workbook from URL instead of reading local file
        const workbook = await fetchRatesheetFromURL();

        // Process cover sheet for contract dates
        const { contractEffectiveDate, contractExpirationDate } = await processCoverSheet(workbook);

        let uswcRates = [];
        let usecRates = [];
        let uswcPreDictRates = [];
        let usecPreDictRates = [];
        let uswcPostDictRates = [];
        let usecPostDictRates = [];

        // Process USWC rates
        if (processMode === 'USWC' || processMode === 'BOTH') {
            const uswcResult = await processUSWCSheet(workbook, contractEffectiveDate, contractExpirationDate);
            uswcRates = uswcResult.processedRates;
            uswcPreDictRates = uswcResult.preDictionaryRates;
            uswcPostDictRates = uswcResult.postDictionaryRates;

            // Write PreDictionaryRates.csv for USWC
            const wsUSWCPre = XLSX.utils.json_to_sheet(uswcPreDictRates);
            const csvUSWCPre = XLSX.utils.sheet_to_csv(wsUSWCPre);
            const csvUSWCPreWithTs = `Timestamp,${timestamp}\n${csvUSWCPre}`;
            require('fs').writeFileSync('PreDictionaryRates.csv', csvUSWCPreWithTs, 'utf8');

            // Write PostDictionaryRates.csv for USWC
            const wsUSWCPost = XLSX.utils.json_to_sheet(uswcPostDictRates);
            const csvUSWCPost = XLSX.utils.sheet_to_csv(wsUSWCPost);
            const csvUSWCPostWithTs = `Timestamp,${timestamp}\n${csvUSWCPost}`;
            require('fs').writeFileSync('PostDictionaryRates.csv', csvUSWCPostWithTs, 'utf8');
        }

        // Process USEC rates
        if (processMode === 'USEC' || processMode === 'BOTH') {
            const usecResult = await processUSECSheet(workbook, contractEffectiveDate, contractExpirationDate);
            usecRates = usecResult.processedRates;
            usecPreDictRates = usecResult.preDictionaryRates;
            usecPostDictRates = usecResult.postDictionaryRates;

            // Write PreDictionaryRates.csv for USEC (append or create)
            const wsUSECPre = XLSX.utils.json_to_sheet(usecPreDictRates);
            const csvUSECPre = XLSX.utils.sheet_to_csv(wsUSECPre);
            const csvUSECPreWithTs = `Timestamp,${timestamp}\n${csvUSECPre}`;
            const preDictFile = 'PreDictionaryRates.csv';
            if (require('fs').existsSync(preDictFile) && processMode === 'BOTH') {
                const existingPreDict = require('fs').readFileSync(preDictFile, 'utf8').split('\n').slice(1).join('\n');
                require('fs').writeFileSync(preDictFile, `Timestamp,${timestamp}\n${existingPreDict}\n${csvUSECPre.split('\n').slice(1).join('\n')}`, 'utf8');
            } else {
                require('fs').writeFileSync(preDictFile, csvUSECPreWithTs, 'utf8');
            }

            // Write PostDictionaryRates.csv for USEC (append or create)
            const wsUSECPost = XLSX.utils.json_to_sheet(usecPostDictRates);
            const csvUSECPost = XLSX.utils.sheet_to_csv(wsUSECPost);
            const csvUSECPostWithTs = `Timestamp,${timestamp}\n${csvUSECPost}`;
            const postDictFile = 'PostDictionaryRates.csv';
            if (require('fs').existsSync(postDictFile) && processMode === 'BOTH') {
                const existingPostDict = require('fs').readFileSync(postDictFile, 'utf8').split('\n').slice(1).join('\n');
                require('fs').writeFileSync(postDictFile, `Timestamp,${timestamp}\n${existingPostDict}\n${csvUSECPost.split('\n').slice(1).join('\n')}`, 'utf8');
            } else {
                require('fs').writeFileSync(postDictFile, csvUSECPostWithTs, 'utf8');
            }
        }

        // Combine rates
        const allRates = [...uswcRates, ...usecRates];

        // Process Service IDs
        await processServiceIDs(workbook);

        // Generate final rates for endpoint
        await createFinalRates(allRates, contractEffectiveDate, contractExpirationDate, processMode);

        console.log(`Processing completed successfully for ${processMode}.`);

    } catch (error) {
        console.error(`Error in main process for ${processMode}:`, error);
        process.exit(1);
    }
}

main();