# QHOF-FAK-contract/parser.js

## Overview

This module is the main parser for QHOF-FAK contract Excel files. It reads, cleans, normalizes, and processes the ratesheet data, including main and feeder rates, and outputs a structured JSON array suitable for downstream processing or integration. It also performs deduplication, enrichment (such as adding container maintenance charges and vessel service names), and validation (via the `tester` utility).

---

## Features

- **Dynamic Header Detection:** Automatically finds the correct header row and maps columns, even with varying formats.
- **Data Cleaning & Normalization:** Handles missing values, splits multi-port entries, and ensures all required fields are present.
- **Deduplication:** Removes duplicate or less competitive rates for the same route.
- **Feeder Rate Integration:** Matches and combines feeder rates with main routes, ensuring comprehensive coverage.
- **Container Maintenance Charges:** Adds relevant surcharges to each route.
- **Validation & Reporting:** Integrates with the `tester` utility to validate and summarize the processed data.
- **Excel & JSON Output:** (Commented out) Capable of exporting processed data for auditing.

---

## Main Export

### `parseQHOFFile(workbook)`

#### **Parameters:**

- `workbook` (`XLSX.WorkBook`): The loaded Excel workbook object.

#### **Returns:**

- `Array<Object>`: Array of formatted rate objects, ready for downstream use.

---

## Key Functions

### 1. `deduplicateArray(array)`

Removes duplicate or less competitive rates for the same route and service, ensuring only the best (lowest) rates are kept.

### 2. `fillMissingValues(array)`

Fills missing "20p" and "45HC" values using "40p" as a base (90% for 20p, 120% for 45HC).

### 3. `findFeederMatches(parsedData, feederRates)`

Finds and combines feeder rates with main routes, creating new entries for routes that require a feeder leg.

### 4. `getHeaderRow(workbook)`

Dynamically detects the header row by searching for key terms, with fallbacks for robustness.

### 5. `getHeaderMap(workbook)`

Builds a mapping from normalized header names to column indices, supporting flexible Excel formats.

### 6. `formatParsedData(parsedData)`

Formats the final output as an array of objects with stringified fields, matching the expected schema.

---

## Processing Flow

1. **Initialize Ports:** Ensures port data is loaded for lookups.
2. **Extract Contract Number:** Scans the first rows for a QHOF contract identifier.
3. **Header Mapping:** Dynamically finds and maps column headers.
4. **Row Parsing:** Iterates through data rows, extracting and normalizing fields.
5. **Multi-Port Handling:** Splits rows with multiple origins into separate entries.
6. **Service Name Enrichment:** Populates the `service` field for each route.
7. **Container Maintenance Charges:** Adds surcharges where applicable.
8. **Filter & Deduplicate:** Removes entries with special container types (SOC, NOR, HAZ) and ensures all required fields are present.
9. **Feeder Rate Integration:** Parses and matches feeder rates, adding new combined routes.
10. **Validation:** Runs the `tester` utility for auditing and reporting.
11. **Final Formatting:** Converts all fields to strings and returns the result.

---

## Example Output

```json
[
  {
    "carrier": "653309",
    "port_destination": "12345",
    "port_origin": "67890",
    "port_discharge": "54321",
    "vf": "2024-06-01",
    "vt": "2024-12-31",
    "transit_time": "",
    "rate_source": "653309",
    "40p": "2000",
    "40hqp": "2100",
    "20p": "1800",
    "45HC": "2400",
    "service": "CMA CGM Service",
    "contract": "33",
    "carrier_contract_number": "QHOF123456"
  }
]
```

---

## Dependencies

- [`xlsx`](https://www.npmjs.com/package/xlsx): Excel file parsing.
- [`../resources/utils.js`]: Utility functions for normalization, port lookups, date conversion, and Excel/JSON output.
- [`../resources/api.service.js`]: API calls for vessel service names and error reporting.
- [`./container_maintenance_parser.js`]: Parses container maintenance charges.
- [`./feeder_parser.js`]: Parses feeder rates.
- [`./tester.js`]: Validation and reporting utility.

---

## Usage Example

```js
import XLSX from "xlsx";
import { parseQHOFFile } from "./QHOF-FAK-contract/parser.js";

const workbook = XLSX.readFile("path/to/QHOF-file.xlsx");
const parsedRates = await parseQHOFFile(workbook);

console.log(parsedRates); // Array of normalized rate objects
```

---

## Notes

- **Error Handling:** Sends error messages via the API service if headers or required fields are missing.
- **Extensible:** Designed to handle changes in Excel format with minimal code changes.
- **Auditing:** Integrates with the `tester` utility for output validation and Excel reporting.
- **Commented Output:** JSON and Excel output lines are present but commented out; enable as needed for debugging or auditing.

---

## File Location

`QHOF-FAK-contract/parser.js`

---

## Summary

This parser is the backbone of the QHOF-FAK contract ratesheet processing pipeline. It robustly handles messy, inconsistent Excel files, enriches and validates the data, and outputs a clean, deduplicated, and feeder-integrated rates array ready for further use or integration.
