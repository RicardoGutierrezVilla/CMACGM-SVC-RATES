// locations.js
const fs = require('fs');
const path = require('path');
const axios = require('axios');

const locationsFile = 'LocationsBettyBlocks.json';
const locationsEndpoint = 'https://primefreight.bettyblocks.com/api/models/companies/records/?view_id=3d419df0045d4a139e5e73902ca2073a&limit=500';
const basicAuthHeader = 'Basic YXBwQHByaW1lZnJlaWdodC5jb206ZmQzZWQ2ZTk4ZDljYzJhMGE2MWJhMzdjZDBmYWU5NjU=';

async function fetchAndSaveLocations() {
    try {
        console.log(`${locationsFile} will be overwritten with fresh data from Betty Blocks.`);
        const response = await axios.get(locationsEndpoint, {
            headers: { 'Authorization': basicAuthHeader },
            responseType: 'json'
        });
        fs.writeFileSync(locationsFile, JSON.stringify(response.data, null, 2), 'utf8');
        console.log(`Locations data written to ${locationsFile}`);
        return response.data;
    } catch (error) {
        console.error("Error fetching locations:", error.response?.data || error.message);
        throw error;
    }
}

module.exports = { fetchAndSaveLocations };