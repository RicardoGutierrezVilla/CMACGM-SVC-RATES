# cma-ratesheet-parser

CMA / CGM Ratesheet Parser

## Initialization

- Uses Node v22.3.0 (npm v10.8.1)
- Run `npm install`

## Packages

- `axios`: For API Calls
- `XLXS`: For interpretting and parsing excel files
  - Loaded from https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz for a more updated _open-sourced_ alternative

## Repository Structure

### `QHOF-FAK-contract`

- All the files for parsing the QHOF contract

### `main.js`

- Entry file for Apify
- Get the file from SFTP
- Runs a contract parser based on file

### `/resources`

- **`api.service.js`**: Contains all API functions
- **`utils.js`**: Contains all utility functions
# CMACGM-SVC-RATES
# CMACGM-SVC-RATES
# CMACGM-SVC-RATES
