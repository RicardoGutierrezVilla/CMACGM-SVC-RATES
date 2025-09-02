import axios from "axios";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { logMatchResult, logAPIFailure } from "./utils.js";
import { get } from "http";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const MAX_RETRIES = 3;
const MAX_RETRY_COUNT = 3;
const RETRY_DELAY_MS = 2000;
const RETRY_DELAY = 2000;

export const BETTY_API_URL =
  "https://primefreight.betty.app/api/runtime/da93364a26fb4eeb9e56351ecec79abb";
export const FCL_MAKE_API_URL =
  "https://hook.us1.make.com/pb3chu1cv36412wo9nzjeupbwo5ucvsw";
export const LCL_MAKE_API_URL =
  "https://hook.us1.make.com/x3i2x9reorn2bigpkvcuqte6fk7g5p2y";
export const MAKE_API_URL =
  "https://hook.us1.make.com/3ewsb0bi54wrow4b0ivp695wrc2n8ijk";
export const MAKE_ERROR_URL =
  "https://hook.us1.make.com/a1v4v6su58bf91xv1u15s2lhg7k8fwnp";

/**
 * @description Handles failure of requests
 * @param {Request} requestPromise
 * @returns Returns the result of the request
 */
export function handleRequest(requestPromise) {
  return requestPromise
    .then((result) => {
      if (result.status === 200) {
        return result;
      } else if (result.status !== 200) {
        throw new Error(`HTTP error! Status: ${result.status}`);
      }
    })
    .catch((error) => {
      console.error("Error in handleRequest:", error);
    });
}
/**
 * @description Gets an array of ports
 * @returns Array of all ports from betty blocks
 */
export function getAllPorts() {
  return axios
    .get(`https://primefreight.bettywebblocks.com/preview?type=ports`)
    .then((result) => {
      let allPort = result?.data?.Ports;
      return allPort;
    });
}

/**
 * @description Function that loads the SCAC Codes
 * @returns Dictionary of initialized SCAC Codes
 */
export async function initializeSCACCodes() {
  const token =
    "Basic YXBwQHByaW1lZnJlaWdodC5jb206ZmQzZWQ2ZTk4ZDljYzJhMGE2MWJhMzdjZDBmYWU5NjU=";
  return axios
    .get(
      `https://primefreight.bettyblocks.com/api/models/companies/records/?view_id=3d419df0045d4a139e5e73902ca2073a&limit=500`,
      {
        headers: {
          Authorization: token,
          "Content-Type": "application/json",
        },
      }
    )
    .then((result) => {
      const records = result?.data?.records || [];
      const dict = {};
      records.forEach((record) => {
        const key = record["ff0a74ce0a164d259adc7b11eeb77334"]?.value;
        const value = record["44a2221218314e1e8359a19a98fe72a1"]?.value;
        if (key && value) {
          dict[key] = value;
        }
      });
      return dict;
    });
}

/**
 * @description This function gets the route from the google sheets
 * @param {Array} routeArray - An array of JSON objects with the port origin id and port discharge id
 * @returns {Promise} - A promise that resolves to the response data.
 */

/**
 *
 * @returns {Promise} - A promise that resolves to the response data.
 * @description This function fetches all ports from the API and returns the data.
 */
export async function getRecordsFromSheet() {
  const SPREADSHEET_ID = "1yBg3JcGlt_Jhnegd-AOEAx83m6xSm5Ee3afZWcMBSNI";
  const GID = "1057237072";
  const CSV_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=csv&gid=${GID}`;

  try {
    const response = await axios.get(CSV_URL);
    const rows = response.data.split("\n").map((row) => row.split(","));

    // Skip header row and map the data
    const records = rows
      .slice(1)
      .filter((row) => row[4] !== undefined && row[4].trim() !== "") // Remove rows with missing service_id
      .map((row) => ({
        origin_id: row[0],
        discharge_id: row[1],
        service_name: row[3],
        service_id: parseInt(row[4].replace(/[^\d]/g, "")),
      }))
      .filter(
        (record) => record.origin_id?.trim() && record.discharge_id?.trim()
      ); // Filter out empty rows

    return records;
  } catch (error) {
    console.error("Error fetching sheet data:", error);
    return [];
  }
}

function normalizePortName(portName) {
  if (!portName) return "";
  // Remove common suffixes and clean the name
  return portName
    .split(/[\n,]+/)[0] // Take first part before comma or newline
    .replace(/\s*,.*$/, "") // Remove everything after comma
    .replace(/\s+(BC|QC|ON|AB|MB|NB|NL|NS|NT|NU|PE|SK|YT)$/i, "") // Remove provinical code
    .trim()
    .toLowerCase();
}

function isPortMatch(sourcePort, targetPort) {
  if (!sourcePort || !targetPort) return false;

  const normalizedSource = normalizePortName(sourcePort);
  const normalizedTarget = normalizePortName(targetPort);

  // First try exact match
  if (normalizedSource === normalizedTarget) return true;

  // Then try word-by-word matching
  const sourceWords = normalizedSource.split(/\s+/);
  const targetWords = normalizedTarget.split(/\s+/);

  // Match on the first word (most significant part of port name)
  if (sourceWords[0] === targetWords[0]) return true;

  // Check if all words from the shorter name exist in the longer name
  const [shorterWords, longerWords] =
    sourceWords.length < targetWords.length
      ? [sourceWords, targetWords]
      : [targetWords, sourceWords];

  return shorterWords.every((word) =>
    longerWords.some((targetWord) => targetWord === word)
  );
}

/**
 * @description This function fetches all vessel service names for a port origin id and port discharge id.
 * It first checks the Google Sheet for matches, then falls back to the API for missing routes. Updates the given array.
 * @param {Array} routesArray - An array of JSON objects with the port origin id and port discharge id and their respective names
 * @returns {null} - Does not return anything, updates the routesArray in place
 */
export async function getVesselServiceName(routesArray) {
  try {
    // First check Google Sheets for matches
    const sheetRecords = await getRecordsFromSheet();

    // Create arrays for found and remaining routes
    const noVesselServiceFound = [];

    // Check each requested route against sheet data
    for (const route of routesArray) {
      if (!route.port_origin_name || !route.port_discharge_name) {
        console.warn("Missing port names in route:", route);
        continue;
      }

      const matchingRecord = sheetRecords.find(
        (record) =>
          isPortMatch(record.origin_id, route.port_origin_name) &&
          isPortMatch(record.discharge_id, route.port_discharge_name)
      );

      if (matchingRecord) {
        // Update the service of the overall JSON object
        route.service = matchingRecord.service_id;
        route.service_name = matchingRecord.service_name;
        // logMatchResult(route, matchingRecord, "sheet");
      } else {
        // No match found in the sheet, add to routes to update via API
        noVesselServiceFound.push(route);
      }
    }

    // If we have routes not found in sheets, fetch them from API
    // and update them using updateGoogleSheets()
    if (noVesselServiceFound.length > 0) {
      for (const route of noVesselServiceFound) {
        const query = {
          query: "mutation { action(id: $action_id input: $input)}",
          variables: {
            action_id: "32c4095339884c2da3149b3a8c68bb11",
            input: {
              discharge_id: route.port_discharge,
              origin_id: route.port_origin,
            },
          },
        };

        const apiResult = await handleRequest(axios.post(BETTY_API_URL, query));
        const apiRoute = apiResult?.data?.data?.action?.results;

        if (!apiRoute) {
          console.error("No results found in API response");
          logAPIFailure(route, apiResult);
        } else if (apiRoute && apiRoute.service_id) {
          // Update the route with API data
          // This will be reflected in our original routesArray
          route.service = apiRoute.service_id || null;
          route.service_name = apiRoute.service_name || null;

          // Finally, update the Google Sheets with the new data via make
          const updateResponse = await updateGoogleSheets({
            origin_id: route.port_origin_name,
            discharge_id: route.port_discharge_name,
            service_name: route.service_name,
            service_id: route.service,
          });

          // logMatchResult(route, apiRoute, "api", !!updateResponse);
        }
      }
    }
  } catch (error) {
    await sendErrorMessage(`Error in getVesselServiceName: ${error}`);
    console.error("Error in getVesselServiceName:", error);
    return;
  }
}

/**
 * @description This function takes a single updated vessel service and updates the google sheets
 * @param {JSON} data - The data containing the origin port, discharge port, service name and Betty Blocks service id
 * @returns {Promise} - returns the response from Make
 */
export function updateGoogleSheets(data) {
  if (!data) {
    console.error("No data provided to updateGoogleSheets");
    return null;
  }
  if (
    !data.origin_id ||
    !data.discharge_id ||
    !data.service_id ||
    !data.service_name
  ) {
    console.error("Missing required fields in data:", data);
    return null;
  }

  try {
    const repsonse = axios.post(MAKE_API_URL, data);
    return repsonse;
  } catch (error) {
    console.error("Error updating Google Sheets:", error);
    return null;
  }
}

/**
 *
 * @description This function sends JSON data to the FCL endpoint.
 * @param {JSON} data
 * @returns {Promise} - A promise that resolves to the response data.
 */
export function sendJSONToFCLEndpoint(data) {
  return handleRequest(axios.post(FCL_MAKE_API_URL, data))
    .then(async (result) => {
      if (result) {
      } else {
      }
    })
    .catch((error) => {
      sendErrorMessage(`Error sending FCL JSON: ${error}`).then(() => {
        console.error("Error sending FCL JSON:", error);
      });
    });
}

/**
 *
 * @description This function sends JSON data to the LCL endpoint.
 * @param {JSON} data
 * @returns {Promise} - A promise that resolves to the response data.
 */
export function sendJSONToLCLEndpoint(data) {
  return handleRequest(axios.post(LCL_MAKE_API_URL, data))
    .then(async (result) => {
      if (result) {
      } else {
      }
    })
    .catch((error) => {
      sendErrorMessage(`Error sending LCL JSON: ${error}`).then(() => {
        console.error("Error sending data:", error);
      });
    });
}

/**
 * Error message to send to an email. Needs to be awaited
 * @param {String} message
 * @returns Promise - this function needs to be awaited
 */
export async function sendErrorMessage(message) {
  return handleRequest(
    axios.post(MAKE_ERROR_URL, {
      "parser-name": "CMA/CGM Parser",
      "error-message": message,
    })
  );
}
