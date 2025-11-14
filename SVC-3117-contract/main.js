const XLSX = require('xlsx');
const axios = require('axios');
const { login } = require('./auth');
const { fetchAndSaveLocations } = require('./locations');
const { processCoverSheet } = require('./coverSheet');
const { processUSWCSheet } = require('./uswcSheet');
const { processUSECSheet } = require('./usecSheet');
const { processISCUSSheet } = require('./iscUsSheet');
const { processISCUSWCSheet } = require('./iscUSWCSheet');
const { processServiceIDs } = require('./serviceIDs');
const { createFinalRates } = require('./finalRates');

async function fetchRatesheetFromURL() {
    // Declaring the URL of the sheet 
    const url = 'https://www.primefreight.com/cma_rates/ratesheet.xlsx';
    console.log(`Fetching ratesheet from: ${url}`);
    
    // Validating the file existence  
    try {
        const response = await axios.get(url, {
            responseType: 'arraybuffer',
            timeout: 30000 
        });
        
        console.log('Ratesheet fetched successfully');
        return XLSX.read(response.data, { type: 'buffer' });
    } catch (error) {
        console.error('Error fetching ratesheet from URL:', error.message);
        throw new Error(`Failed to fetch ratesheet from ${url}: ${error.message}`);
    }
}

async function main() {
    // Configuration based on contract type 
    const processMode = 'ISCUSWC'; // 'USWC', 'USEC', 'ISCUS','ISCUSWC'  or 'BOTH'
    const timestamp = new Date().toISOString();

    try {
        console.log(`Starting rate sheet processing for ${processMode}...`);
        
        // BettyBlocks Authentication
        await login();

        // Fetch and save locations from Bettyblocks 
        await fetchAndSaveLocations();

        // Fetch the workbook from URL 
        const workbook = await fetchRatesheetFromURL();

        // Process cover sheet for contract dates validity and expiration
        const { contractEffectiveDate, contractExpirationDate } = await processCoverSheet(workbook);

        // Initialize arrays to hold rates 
        let uswcRates = [];
        let usecRates = [];
        let iscusRates = [];
        let iscuswcRates = [];
    
        // Pre and Post Dictionary rates fo debugging and validation
        let uswcPreDictRates = [];
        let usecPreDictRates = [];
        let uswcPostDictRates = [];
        let usecPostDictRates = [];
        let iscusPreDictRates = [];
        let iscusPostDictRates = [];
        let iscuswcPreDictRates = [];
        let iscuswcPostDictRates = [];
        let uswcSurchargeDescriptions = [];

        // Processing USWC rates
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
            require('fs').writeFileSync('PreDictionaryRates.csv', csvUSWCPreWithTs, 'utf8');

            // Write PostDictionaryRates.csv for USWC
            const wsUSWCPost = XLSX.utils.json_to_sheet(uswcPostDictRates);
            const csvUSWCPost = XLSX.utils.sheet_to_csv(wsUSWCPost);
            const csvUSWCPostWithTs = `Timestamp,${timestamp}\n${csvUSWCPost}`;
            require('fs').writeFileSync('PostDictionaryRates.csv', csvUSWCPostWithTs, 'utf8');
        }

        // Processing USEC rates
        if (processMode === 'USEC' || processMode === 'BOTH') {
            const usecResult = await processUSECSheet(workbook, contractEffectiveDate, contractExpirationDate);
            usecRates = usecResult.processedRates;
            usecPreDictRates = usecResult.preDictionaryRates;
            usecPostDictRates = usecResult.postDictionaryRates;
            // Merge/assign surcharge descriptions from USEC
            const usecSurcharges = usecResult.uswcSurchargeDescriptions || [];
            if (usecSurcharges.length > 0) {
                const merged = Array.from(new Set([...(uswcSurchargeDescriptions || []), ...usecSurcharges]));
                uswcSurchargeDescriptions = merged;
            }

            // Write PreDictionaryRates.csv for USEC (overwrite)
            const wsUSECPre = XLSX.utils.json_to_sheet(usecPreDictRates);
            const csvUSECPre = XLSX.utils.sheet_to_csv(wsUSECPre);
            const csvUSECPreWithTs = `Timestamp,${timestamp}\n${csvUSECPre}`;
            require('fs').writeFileSync('PreDictionaryRates.csv', csvUSECPreWithTs, 'utf8');

            // Write PostDictionaryRates.csv for USEC (overwrite)
            const wsUSECPost = XLSX.utils.json_to_sheet(usecPostDictRates);
            const csvUSECPost = XLSX.utils.sheet_to_csv(wsUSECPost);
            const csvUSECPostWithTs = `Timestamp,${timestamp}\n${csvUSECPost}`;
            require('fs').writeFileSync('PostDictionaryRates.csv', csvUSECPostWithTs, 'utf8');
        }

        // Processing ISC-US rates 
        if (processMode === 'ISCUS' || processMode === 'BOTH') {
            const iscusResult = await processISCUSSheet(workbook, contractEffectiveDate, contractExpirationDate);
            iscusRates = iscusResult.processedRates || [];
            iscusPreDictRates = iscusResult.preDictionaryRates || [];
            iscusPostDictRates = iscusResult.postDictionaryRates || [];

            // Merge surcharge descriptions coming from ISC-US sheet
            const iscusSurcharges = iscusResult.uswcSurchargeDescriptions || [];
            if (iscusSurcharges.length > 0) {
                const merged = Array.from(new Set([...(uswcSurchargeDescriptions || []), ...iscusSurcharges]));
                uswcSurchargeDescriptions = merged;
            }

            // Write PreDictionaryRates.csv for ISC-US (overwrite)
            if (iscusPreDictRates.length > 0) {
                const wsISCPre = XLSX.utils.json_to_sheet(iscusPreDictRates);
                const csvISCPre = XLSX.utils.sheet_to_csv(wsISCPre);
                const csvISCPreWithTs = `Timestamp,${timestamp}\n${csvISCPre}`;
                require('fs').writeFileSync('PreDictionaryRates.csv', csvISCPreWithTs, 'utf8');
            } else {
                console.log('ISC-US preDictionary has 0 rows; writing blank PreDictionaryRates.csv');
                require('fs').writeFileSync('PreDictionaryRates.csv', `Timestamp,${timestamp}\n`, 'utf8');
            }

            // Write PostDictionaryRates.csv for ISC-US (overwrite)
            if (iscusPostDictRates.length > 0) {
                const wsISCPost = XLSX.utils.json_to_sheet(iscusPostDictRates);
                const csvISCPost = XLSX.utils.sheet_to_csv(wsISCPost);
                const csvISCPostWithTs = `Timestamp,${timestamp}\n${csvISCPost}`;
                require('fs').writeFileSync('PostDictionaryRates.csv', csvISCPostWithTs, 'utf8');
            } else {
                console.log('ISC-US postDictionary has 0 rows; writing blank PostDictionaryRates.csv');
                require('fs').writeFileSync('PostDictionaryRates.csv', `Timestamp,${timestamp}\n`, 'utf8');
            }
        }

         // Processing ISCUSWC rates 
        if (processMode === 'ISCUSWC' || processMode === 'BOTH') {
            const iscuswcResult = await processISCUSWCSheet(workbook, contractEffectiveDate, contractExpirationDate);
            iscuswcRates = iscuswcResult.processedRates || [];
            iscuswcPreDictRates = iscuswcResult.preDictionaryRates || [];
            iscuswcPostDictRates = iscuswcResult.postDictionaryRates || [];

            // Merge surcharge descriptions coming from ISC-USWC sheet
            const iscuswcSurcharges = iscuswcResult.uswcSurchargeDescriptions || [];
            if (iscuswcSurcharges.length > 0) {
                const merged = Array.from(new Set([...(uswcSurchargeDescriptions || []), ...iscuswcSurcharges]));
                uswcSurchargeDescriptions = merged;
            }

            // Write PreDictionaryRates.csv for ISC-USWC (overwrite)
            if (iscuswcPreDictRates.length > 0) {
                const wsPre = XLSX.utils.json_to_sheet(iscuswcPreDictRates);
                const csvPre = XLSX.utils.sheet_to_csv(wsPre);
                const csvPreWithTs = `Timestamp,${timestamp}\n${csvPre}`;
                require('fs').writeFileSync('PreDictionaryRates.csv', csvPreWithTs, 'utf8');
            } else {
                require('fs').writeFileSync('PreDictionaryRates.csv', `Timestamp,${timestamp}\n`, 'utf8');
            }

            // Write PostDictionaryRates.csv for ISC-USWC (overwrite)
            if (iscuswcPostDictRates.length > 0) {
                const wsPost = XLSX.utils.json_to_sheet(iscuswcPostDictRates);
                const csvPost = XLSX.utils.sheet_to_csv(wsPost);
                const csvPostWithTs = `Timestamp,${timestamp}\n${csvPost}`;
                require('fs').writeFileSync('PostDictionaryRates.csv', csvPostWithTs, 'utf8');
            } else {
                require('fs').writeFileSync('PostDictionaryRates.csv', `Timestamp,${timestamp}\n`, 'utf8');
            }
        }

        // Combining rates based on selected mode 
        let allRates = [];
        if (processMode === 'USWC') allRates = [...uswcRates];
        else if (processMode === 'USEC') allRates = [...usecRates];
        else if (processMode === 'ISCUS') allRates = [...iscusRates];
        else if (processMode === 'ISCUSWC') allRates = [...iscuswcRates];
        else /* BOTH */ allRates = [...uswcRates, ...usecRates, ...iscusRates, ...iscuswcRates];

        // Processing Service IDs
        await processServiceIDs(workbook);

        // Generating and formatting final rates for endpoint
        await createFinalRates(
            allRates,
            contractEffectiveDate,
            contractExpirationDate,
            processMode,
            uswcSurchargeDescriptions
        );
        console.log(`Processing completed successfully for ${processMode}.`);

    } catch (error) {
        console.error(`Error in main process for ${processMode}:`, error);
        process.exit(1);
    }
}

main();
