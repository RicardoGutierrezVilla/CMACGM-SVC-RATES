import { getAllPorts } from "./api.service.js";
import { appendFileSync, existsSync, mkdirSync } from "fs";
import * as fs from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import XLSX from "xlsx";
import { tester } from "../QHOF-FAK-contract/tester.js";

let portsCache = null;

/**
 * Converts an Excel date to a JavaScript date.
 * @param {number} excelDate - The Excel date number.
 * @returns {string|null} - The formatted date string in YYYY-MM-DD format or null if the input is invalid.
 * @description This function converts an Excel date number to a JavaScript date string in the format YYYY-MM-DD.
 **/
export function excelDateToJSDate(excelDate) {
  if (!excelDate || isNaN(excelDate)) return null;

  // Excel's epoch starts from 1900-01-01
  const date = new Date((excelDate - 25569) * 86400 * 1000);
  // Format as YYYY-MM-DD
  return date.toISOString().split("T")[0];
}

/**
 * Initialize ports data - call this once at the start
 * @returns {Promise} - A promise that resolves to the ports data.
 * @description This function initializes the ports data by fetching it from the API.
 * It caches the result to avoid multiple API calls.
 * If the data is already cached, it returns the cached data.
 */
export async function initializePorts() {
  if (!portsCache) {
    portsCache = await getAllPorts();
  }
  return portsCache;
}

/**
 *
 * @param {string} str
 * @returns A normalized string without punctuation, extra spaces and lowercase.
 * @description This function normalizes a string by converting it to lowercase,
 * removing punctuation, and collapsing whitespace.
 */
export const normalize = (str) =>
  str
    .toLowerCase() // Convert to lowercase
    .replace(/[\(\)\[\],.]/g, "") // Remove punctuation and parentheses
    .replace(/\s+/g, " ") // Collapse whitespace
    .trim();

/**
 * Fetches database ports and matches them with the provided port name.
 * It uses a predefined set of regex patterns to find the best match.
 *
 * @param {string} portName
 * @returns {int} - The ID of the port.
 * @description This function retrieves the port ID based on the provided port name.
 */
export async function getPortId(portName) {
  if (!portName) return null;

  if (!portsCache) {
    await initializePorts();
  }

  // Split input by newlines and commas
  const portParts = portName
    .split(/[\n]+/)
    .map((p) => p.trim())
    .filter(Boolean);

  // Use the last part if it contains a location identifier (BC, ON, QC, etc.)
  // Otherwise use the first non-empty part
  const cleanPortName =
    portParts.find((p) => /[A-Z]{2}$/.test(p)) || portParts[0];

  // If no valid port name is found, return null
  if (!cleanPortName) return null;

  // Create a more strict regex pattern
  // This pattern matches the cleaned port name exactly, ignoring case
  // and allowing for variations in whitespace
  const pattern = new RegExp(`^${cleanPortName.split(",")[0].trim()}$`, "i");

  const portId = portsCache.find((p) => {
    // Try exact match first
    if (pattern.test(p.Name)) return true;

    // Then try partial match but with stricter rules
    const portNameWords = cleanPortName.toLowerCase().split(/\s+/);
    const dbPortWords = p.Name.toLowerCase().split(/\s+/);

    // If database port has only one word, match only the first word
    if (dbPortWords.length === 1) {
      return portNameWords[0] === dbPortWords[0];
    }

    // All words from input must exist in database port name
    return portNameWords.every((word) =>
      dbPortWords.some((dbWord) => normalize(dbWord) === normalize(word))
    );
  });

  return portId ? portId.id : null;
}

/*

TEST FUNCTIONS -- Useful in debugging purposes

*/
// Get current file path in ES modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/**
 *
 * @description Logs the input and result to a file for debugging purposes.
 * @param {string} portName
 * @param {JSON} result
 */
function logToFile(portName, result) {
  const logPath = join(__dirname, "../output/port-name-matches.log");
  const timestamp = new Date().toISOString();
  const logEntry = `${timestamp} | Input: "${portName}" | Result: ${JSON.stringify(
    result
  )}\n`;

  try {
    appendFileSync(logPath, logEntry);
  } catch (error) {
    console.error("Error writing to debug log:", error);
  }
}

/**
 * Logs null matches to a file.
 *
 * @param {string} portName
 * @param {JSON} result
 */
function logNullMatches(portName, result) {
  const logPath = join(__dirname, "../output/port-null-matches.log");
  const timestamp = new Date().toISOString();
  const logEntry = `${timestamp} | Input: "${portName}" | Result: ${JSON.stringify(
    result
  )}\n`;

  try {
    appendFileSync(logPath, logEntry);
  } catch (error) {
    console.error("Error writing to null matches log:", error);
  }
}

/**
 * @description Logs matches found in either sheet or API, and their update status
 * @param {Object} route - The route being matched
 * @param {Object} matchResult - The match result (from sheet or API)
 * @param {string} source - The source of the match ('sheet' or 'api')
 * @param {boolean} updateSuccess - Whether the Google Sheets update was successful (only for API matches)
 */
export function logMatchResult(
  route,
  matchResult,
  source,
  updateSuccess = null
) {
  const timestamp = new Date().toISOString();
  let logMessage = `${timestamp} | Match found in ${source.toUpperCase()}:\n`;

  const matchDetails = {
    origin_id:
      source === "sheet" ? matchResult.origin_id : route.port_origin_name,
    discharge_id:
      source === "sheet" ? matchResult.discharge_id : route.port_discharge_name,
    service_name: matchResult.service_name,
    service_id: matchResult.service_id,
  };

  logMessage += `${JSON.stringify(matchDetails, null, 2)}\n`;

  if (source === "api" && updateSuccess !== null) {
    logMessage += `Google Sheets update: ${
      updateSuccess ? "SUCCESS" : "FAILED"
    }\n`;
  }

  logMessage += "----------------------------------------\n";

  // File output
  const logPath = join(__dirname, "../output/vessel-service-matches.log");
  try {
    appendFileSync(logPath, logMessage);
  } catch (error) {
    console.error("Error writing to match log:", error);
  }
}

/**
 * @description Logs vessel service matching attempts for debugging
 * @param {Object} route - The route being matched
 * @param {Object} matchResult - The result of the matching attempt
 */
function logVesselServiceMatch(route, matchResult) {
  const logPath = join(__dirname, "vessel-service-matches.log");
  const timestamp = new Date().toISOString();
  const logEntry = `${timestamp} | Route: Origin="${
    route.port_origin_name
  }", Discharge="${
    route.port_discharge_name
  }" | Match Found: ${!!matchResult} | Result: ${JSON.stringify(
    matchResult
  )}\n`;

  try {
    appendFileSync(logPath, logEntry);
  } catch (error) {
    console.error("Error writing to vessel service matches log:", error);
  }
}

/**
 * @description Logs failed API attempts for debugging
 * @param {Object} route - The route that failed to match
 * @param {Object} apiResponse - The raw API response
 */
export function logAPIFailure(route, apiResponse) {
  const logPath = join(__dirname, "api-failures.log");
  const timestamp = new Date().toISOString();

  let logMessage = `${timestamp} | API MATCH FAILED\n`;
  logMessage += `Route Details:\n`;
  logMessage += `  Origin: ${route.port_origin_name} (ID: ${route.port_origin})\n`;
  logMessage += `  Discharge: ${route.port_discharge_name} (ID: ${route.port_discharge})\n`;
  logMessage += `API Response:\n${JSON.stringify(
    apiResponse?.data || {},
    null,
    2
  )}\n`;
  logMessage += "----------------------------------------\n";

  try {
    appendFileSync(logPath, logMessage);
  } catch (error) {
    console.error("Error writing to API failures log:", error);
  }
}

/**
 * Log all port names
 */
function logPortNames() {
  const logPath = join(__dirname, "port-names.log");
  const timestamp = new Date().toISOString();
  const logEntry = `${timestamp} | Port Names: ${JSON.stringify(portsCache)}\n`;

  try {
    appendFileSync(logPath, logEntry);
  } catch (error) {
    console.error("Error writing to port names log:", error);
  }
}

/**
 * Writes a JSON array to an Excel file.
 * @param {Array} jsonData - The array of objects to write.
 * @param {string} filePath - The output Excel file path.
 */
export function writeJSONToExcel(overallData, feederRates, mainRates) {
  XLSX.set_fs(fs);
  if (!Array.isArray(overallData) || overallData.length === 0) {
    throw new Error("Input data must be a non-empty array.");
  }

  if (!Array.isArray(feederRates) || feederRates.length === 0) {
    throw new Error("Feeder rates data must be a non-empty array.");
  }

  if (!Array.isArray(mainRates) || mainRates.length === 0) {
    throw new Error("Main rates data must be a non-empty array.");
  }

  // Convert JSON to worksheet
  const worksheet = XLSX.utils.json_to_sheet(overallData);
  const feederRatesSheet = XLSX.utils.json_to_sheet(feederRates);
  const mainRatesSheet = XLSX.utils.json_to_sheet(mainRates);

  // Create a new workbook and append the worksheet
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Overall Data");
  XLSX.utils.book_append_sheet(workbook, feederRatesSheet, "Feeder Rates");
  XLSX.utils.book_append_sheet(workbook, mainRatesSheet, "Main Rates");
  // Write workbook to file

  const outputDir = join(__dirname, "../output");
  if (!existsSync(outputDir)) {
    mkdirSync(outputDir, { recursive: true });
  }

  const outputPath = join(outputDir, "output.xlsx");
  XLSX.writeFile(workbook, outputPath);
}
