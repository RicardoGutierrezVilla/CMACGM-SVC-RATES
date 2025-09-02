# /resources Directory

This folder contains shared modules used by the parsers and other parts of the application. The files provide API helpers and utility functions for working with port data, Excel files, and logging.

## `api.service.js`

Utilities for fetching data from Betty Blocks, sending information to Make hooks, and retrieving vessel service names.

### `handleRequest(requestPromise)`
Handles an axios request promise and standardizes error checking.

**Parameters**
- `requestPromise` (Promise): The axios request to resolve.

**Logic**
- Waits for the request to resolve.
- If the response has status `200`, returns the result.
- Throws an error for non-`200` responses and logs errors caught during execution.

### `getAllPorts()`
Fetches the list of all ports from Betty Blocks.

**Parameters**
- none

**Logic**
- Performs a GET request to `primefreight.bettywebblocks.com`.
- Returns the `Ports` array from the response.

### `initializeSCACCodes()`
Retrieves a table of SCAC codes from Betty Blocks and converts it into a lookup dictionary.

**Parameters**
- none

**Logic**
- Sends an authenticated GET request to the companies records endpoint.
- Iterates over records and builds a dictionary of code to value pairs.
- Returns the constructed dictionary.

### `getRecordsFromSheet()`
Loads port/service records from a Google Sheet in CSV format.

**Parameters**
- none

**Logic**
- Downloads the CSV, splits rows, and skips the header.
- Builds an array of objects with origin/discharge IDs, service name, and service ID.
- Filters out empty rows and returns the array.

### `normalizePortName(portName)`
Internal helper that cleans up port names for comparison.

**Parameters**
- `portName` (string): The raw port name.

**Logic**
- Removes extra text (commas, province codes, line breaks) and lowercases the result.

### `isPortMatch(sourcePort, targetPort)`
Checks if two port names refer to the same location.

**Parameters**
- `sourcePort` (string): The first port name.
- `targetPort` (string): The second port name.

**Logic**
- Normalizes both names using `normalizePortName`.
- Checks for exact match, then compares the first word, then ensures all words from the shorter name appear in the longer one.

### `getVesselServiceName(routesArray)`
Finds the vessel service for each route and updates the provided array.

**Parameters**
- `routesArray` (Array): Objects containing origin and discharge port data.

**Logic**
- Looks up each route in the Google Sheet data.
- For unmatched routes, queries the Betty Blocks API.
- Updates the route entries with service ID and name.
- New matches are pushed back to Google Sheets via `updateGoogleSheets`.

### `updateGoogleSheets(data)`
Posts new service data back to the Make integration.

**Parameters**
- `data` (Object): Contains `origin_id`, `discharge_id`, `service_name`, and `service_id`.

**Logic**
- Validates required fields.
- Sends the data via POST to `MAKE_API_URL` and returns the response.

### `sendJSONToFCLEndpoint(data)`
Sends JSON payloads to the FCL Make endpoint.

**Parameters**
- `data` (Object): The payload to send.

**Logic**
- Uses `handleRequest` to POST to `FCL_MAKE_API_URL`.
- Logs and forwards any errors using `sendErrorMessage`.

### `sendJSONToLCLEndpoint(data)`
Sends JSON payloads to the LCL Make endpoint.

**Parameters**
- `data` (Object): The payload to send.

**Logic**
- Uses `handleRequest` to POST to `LCL_MAKE_API_URL`.
- Logs and forwards any errors using `sendErrorMessage`.

### `sendErrorMessage(message)`
Reports an error to the Make error hook.

**Parameters**
- `message` (string): Error text to forward.

**Logic**
- POSTs the message using `handleRequest` to `MAKE_ERROR_URL`.
- Designed to be awaited so calling code can ensure delivery.

---

## `utils.js`

Helper functions for port lookups, normalization, and Excel output.

### `excelDateToJSDate(excelDate)`
Converts an Excel serial date to a human-readable string.

**Parameters**
- `excelDate` (number): Excel date value.

**Logic**
- Calculates the JavaScript date by adjusting for the Excel epoch.
- Returns the date in `YYYY-MM-DD` format or `null` for invalid input.

### `initializePorts()`
Caches and returns the list of ports.

**Parameters**
- none

**Logic**
- Calls `getAllPorts` if the cache is empty.
- Stores the result in `portsCache` and returns it on subsequent calls.

### `normalize(str)`
Simplifies a string for comparison.

**Parameters**
- `str` (string): The text to normalize.

**Logic**
- Lowercases the text, removes punctuation, collapses whitespace, and trims.

### `getPortId(portName)`
Attempts to resolve a port name to its database ID.

**Parameters**
- `portName` (string): Name to search for.

**Logic**
- Initializes ports if not loaded.
- Cleans the input name and builds a strict regex.
- Searches `portsCache` for an exact or partial match.
- Returns the matched port ID or `null`.

### `logToFile(portName, result)`
Writes debug information about port matching to a log file.

**Parameters**
- `portName` (string): Input name.
- `result` (Object): Matching result.

**Logic**
- Appends a timestamped line to `port-name-matches.log`.

### `logNullMatches(portName, result)`
Logs cases where no port match was found.

**Parameters**
- `portName` (string): Input name.
- `result` (Object): Data attempted.

**Logic**
- Appends a timestamped entry to `port-null-matches.log`.

### `logMatchResult(route, matchResult, source, updateSuccess)`
Records matches from either the sheet or API.

**Parameters**
- `route` (Object): The route being processed.
- `matchResult` (Object): Found service information.
- `source` (string): Either `"sheet"` or `"api"`.
- `updateSuccess` (boolean, optional): Result of updating Google Sheets when source is API.

**Logic**
- Formats a log entry with the match details.
- Writes the entry to `vessel-service-matches.log`.

### `logVesselServiceMatch(route, matchResult)`
Internal helper for debugging vessel service lookups.

**Parameters**
- `route` (Object): The route inspected.
- `matchResult` (Object): Result information.

**Logic**
- Appends a timestamped entry to `vessel-service-matches.log`.

### `logAPIFailure(route, apiResponse)`
Logs details of failed API lookups.

**Parameters**
- `route` (Object): Route information.
- `apiResponse` (Object): Raw response from the API.

**Logic**
- Writes the route data and API response to `api-failures.log`.

### `logPortNames()`
Dumps all cached port names to a log for review.

**Parameters**
- none

**Logic**
- Appends the current list of ports to `port-names.log`.

### `writeJSONToExcel(overallData, feederRates, mainRates)`
Saves parsed rate data into an Excel workbook with multiple sheets.

**Parameters**
- `overallData` (Array): The aggregated rate objects.
- `feederRates` (Array): Feeder rate entries.
- `mainRates` (Array): Main rate entries.

**Logic**
- Converts each array to a worksheet using `xlsx` utilities.
- Creates an output directory if missing and writes `output.xlsx` containing three sheets.

---

File Location: `/resources/README.md`
