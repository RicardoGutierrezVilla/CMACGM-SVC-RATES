# cma-ratesheet-parser

CMA/CGM Ratesheet Parser packaged as an Apify Actor.

This actor downloads a CMA/CGM ratesheet Excel file, detects the contract type from the workbook, forwards processing to the correct sub-module, and optionally pushes the resulting rates to the Apify dataset for downstream use.

## How It Works

1) Init and Input
- Initializes the Apify SDK and reads input as defined in `INPUT_SCHEMA.json`.
- Optional delay (`delayBeforeMs`) before starting.

2) Download
- Downloads the ratesheet from `ratesheetUrl` (defaults to `https://www.primefreight.com/cma_rates/ratesheet.xlsx`).

3) Detect Contract Type
- Reads the workbook and checks the Cover sheet for a Service Contract indicator.
- If it matches SVC (3117/3118), processing continues; otherwise, an error is logged and sent to the error webhook.

4) Route to Sub-Processor
- Mode selection: 
  - `AUTO` detects whether to use the 3117 Contract processor or the Feeder processor.
  - `SVC_3117_CONTRACT` forces `SVC-3117-contract`.
  - `SVC_3117_FEEDER` forces `SVC-3117-Feeder`.
- The selected module runs and generates output files (e.g., `FinalRatesToEndpoint.json`).

5) Output
- If `pushResults` is true, the actor looks for a `FinalRatesToEndpoint.json` and pushes its content to the default Apify dataset.
- Temporary files are cleaned up.

## Inputs

Provided via `INPUT_SCHEMA.json` and the Apify UI:

- `ratesheetUrl` (string): Optional URL for the CMA/CGM Excel file. Default: `https://www.primefreight.com/cma_rates/ratesheet.xlsx`.
- `mode` (string): One of `AUTO`, `SVC_3117_CONTRACT`, `SVC_3117_FEEDER`. Default: `AUTO`.
- `delayBeforeMs` (integer): Optional delay before processing, in milliseconds. Default: `0`.
- `pushResults` (boolean): Push results to Apify dataset if found. Default: `true`.

Example JSON input:

```
{
  "ratesheetUrl": "https://www.primefreight.com/cma_rates/ratesheet.xlsx",
  "mode": "AUTO",
  "delayBeforeMs": 0,
  "pushResults": true
}
```

## Outputs

- Default dataset items (Apify Dataset): Contents of `FinalRatesToEndpoint.json` when present.
- Local artifacts produced by the sub-modules (for debugging/inspection when running locally):
  - `FinalRatesToEndpoint.json`
  - `PreDictionaryRates.csv`, `PostDictionaryRates.csv`
  - Other generated CSV/JSON files inside `SVC-3117-contract` or `SVC-3117-Feeder`.

## Run on Apify

- Deploy this repository as an Apify Actor.
- Provide input via the actor UI or API using the schema above.
- View results in the Run’s default dataset.

## Run Locally

Prerequisites:
- Node.js 20+ (Apify base image runs Node 20). Repo `package.json` is ESM (`"type": "module"`).

Install and run:
- `npm install`
- `npm start`

With Apify CLI:
- `apify run -p` to use the input UI
- `apify run -i '{"mode":"AUTO","pushResults":true}'`

## Repository Structure

- `index.js`: Actor entry (downloads workbook, detects, forwards to sub-processor, optionally pushes dataset items).
- `INPUT_SCHEMA.json`: Apify input schema for the actor UI and API.
- `resources/`
  - `api.service.js`: External API calls (Make hooks, Betty Blocks endpoints).
  - `utils.js`: Helper utilities (ports, date conversions, logging helpers).
- `SVC-3117-contract/`: Contract flow for 3117; generates final rates and CSVs.
- `SVC-3117-Feeder/`: Feeder flow; generates final rates and CSVs.
- `.actor/actor.json`: Actor metadata (build tag, Dockerfile).
- `Dockerfile`: Uses `apify/actor-node:20` to run the actor in the cloud.

## Error Handling

- Errors are logged to console and sent via `sendErrorMessage` (Make webhook) for visibility.
- If the contract type cannot be detected, the actor logs an error and exits after cleanup.
