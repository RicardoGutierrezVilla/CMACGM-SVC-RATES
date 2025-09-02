import {
  excelDateToJSDate,
  getPortId,
  initializePorts,
  normalize,
} from "../resources/utils.js";
import {
  getVesselServiceName,
  sendErrorMessage,
} from "../resources/api.service.js";
import { parseContainerMaintenanceCharges } from "./container_maintenance_parser.js";
import { parseFeederRates } from "./feeder_parser.js";
import { writeJSONToExcel } from "../resources/utils.js";
import { tester } from "./tester.js";

import XLSX from "xlsx";

/**
 * Deduplicates by getting either checking if all container prices are equal, or if the same route has a higher price
 * @param {Array} array
 */
function deduplicateArray(array) {
  for (const item of array) {
    for (const secondItem of array) {
      if (item !== secondItem) {
        if (
          item.service === secondItem.service &&
          item.port_origin === secondItem.port_origin &&
          item.port_discharge === secondItem.port_discharge &&
          item.port_destination === secondItem.port_destination &&
          item["20p"] === secondItem["20p"] &&
          item["40p"] === secondItem["40p"] &&
          item["40hpq"] === secondItem["40hpq"] &&
          item["45HC"] === secondItem["45HC"]
        ) {
          // If they are duplicates, remove the second one
          array.splice(array.indexOf(secondItem), 1);
        }
        // Else if service, origin, discharge, destination are all equal, then take the lower rate
        else if (
          item.service === secondItem.service &&
          item.port_origin === secondItem.port_origin &&
          item.port_discharge === secondItem.port_discharge &&
          item.port_destination === secondItem.port_destination
        ) {
          if (
            item["20p"] >= secondItem["20p"] &&
            item["40p"] >= secondItem["40p"] &&
            item["40hpq"] >= secondItem["40hpq"] &&
            item["45HC"] >= secondItem["45HC"]
          ) {
            // The first item is the cheaper rate
            array.splice(array.indexOf(item), 1);
          } else if (
            secondItem["20p"] >= item["20p"] &&
            secondItem["40p"] >= item["40p"] &&
            secondItem["40hpq"] >= item["40hpq"] &&
            secondItem["45HC"] >= item["45HC"]
          ) {
            // The second item is the cheaper rate
            array.splice(array.indexOf(secondItem), 1);
          }
        }
      }
    }
  }
}

/**
 * Populates "20p" and "45HC" fields if they are missing by using "40p"
 * @param {Array} array
 */
function fillMissingValues(array) {
  for (const item of array) {
    if (!item["20p"] || item["20p"] === null || item["20p"] === 0) {
      item["20p"] = 0.9 * item["40p"];
    }
    if (!item["45HC"] || item["45HC"] === null || item["45HC"] === 0) {
      item["45HC"] = 1.2 * item["40p"];
    }
  }
}

/**
 * Finds matches between feederRates and parsedData based on:
 * 1. main_pol_id === port_origin
 * 2. eq matches one of the keys [20p, 40p, 40hpq] in parsedData
 * @param {Array} parsedData - The parsed data from the Excel file
 * @param {Array} feederRates - The feeder rates data
 * @param {Set} uniqueOriginPortIds - A set of unique port_origin IDs
 * @description This function finds matches between feeder rates and parsed data
 * based on the main_pol_id and port_origin. It combines the data from both
 * sources into a single object and returns an array of matches. It will not add a feeder rate if the origin port
 * is already in the parsedData.
 * @returns an array of matches: { feederRate, parsedEntry }
 */
function findFeederMatches(parsedData, feederRates) {
  const matches = [];

  // Get a set of unique port_origin IDs
  const uniqueOriginPortIds = new Set(
    parsedData.map((entry) => entry.port_origin)
  );

  for (const feederRate of feederRates) {
    if (!feederRate.main_pol_id) continue;
    if (uniqueOriginPortIds.has(feederRate.port_origin_id)) {
      // Skip if the port_origin is already in parsedData
      // This skips the feeder rate if the port_origin is already in parsedData
      continue;
    }
    // Step 1: Find parsedData entries where port_origin matches main_pol_id
    const candidates = parsedData.filter(
      (entry) => entry.port_origin === feederRate.main_pol_id
    );

    // Step 2: Combine the matches into a single object
    for (const candidate of candidates) {
      // Create a new object to store the combined data
      // Reassign the port_origin and port_origin_name from feederRate
      // to the candidate entry
      const combinedEntry = {
        ...candidate,
        port_origin: feederRate.port_origin_id,
        port_origin_name: feederRate.port_origin_name,
        "20p": (candidate["20p"] || 0) + (feederRate["20p"] || 0),
        "40p": (candidate["40p"] || 0) + (feederRate["40p"] || 0),
        "40hpq": (candidate["40hpq"] || 0) + (feederRate["40hpq"] || 0),
        "45HC": (candidate["45HC"] || 0) + (feederRate["45HC"] || 0),
      };
      matches.push(combinedEntry);
    }
  }

  return matches;
}

/**
 *
 * @description This function finds the header row in the Excel sheet. It first looks for specific keywords
 * in the first few rows. If not found, it defaults to a specific row (6th row).
 * @param {workbook} workbook
 */
async function getHeaderRow(workbook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // 1. Find the header row
  let headerRowIndex = -1;
  for (let i = 0; i < jsonData.length; i++) {
    const row = jsonData[i];
    for (let j = 0; j < row.length; j++) {
      const cell = row[j];
      if (
        typeof cell === "string" &&
        (cell.includes("SOC") ||
          cell.includes("NOR") ||
          cell.includes("HAZ") ||
          normalize(cell).includes("load") ||
          normalize(cell).includes("discharge") ||
          normalize(cell).includes("delivery"))
      ) {
        headerRowIndex = i;
        break;
      }
    }
  }

  if (headerRowIndex === -1) {
    await sendErrorMessage("1ST priority for finding the header row not found");

    for (let i = 0; i < jsonData.length; i++) {
      const row = jsonData[i];
      for (let j = 0; j < row.length; j++) {
        const cell = row[j];
        if (
          typeof cell === "string" &&
          (cell.includes("DG") || cell.includes("NREF"))
        ) {
          headerRowIndex = i;
          break;
        }
      }
    }
  }

  if (headerRowIndex === -1) {
    // Use default row and send error
    await sendErrorMessage("Header Row Not found - Used default row 5");
    headerRowIndex = 5;
  }

  return headerRowIndex;
}

/**
 *
 * @param {workbook} workbook - The workbook object containing the Excel data
 * @returns {Map} - A map of [Header, Header Column Index] pairs
 * @description This function returns a map of the header names and their respective column indices
 * for the main sheet
 */
async function getHeaderMap(workbook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // Get the header row
  const headerRowIndex = await getHeaderRow(workbook);
  const headerRow = jsonData[headerRowIndex];

  // Create a map of header names and their respective column indices
  const headerMap = new Map();
  for (let i = 0; i < headerRow.length; i++) {
    let header = headerRow[i];
    // check if the header is a string
    if (typeof header !== "string") continue;

    header = normalize(header);

    // First load port
    if (header.includes("load") || header.includes("origin")) {
      headerMap.set("Load", i);
    }

    if (header.includes("discharge") || header.includes("unloading")) {
      headerMap.set("Discharge", i);
    }

    if (
      header.includes("delivery") ||
      header.includes("destination") ||
      header.includes("arrival")
    ) {
      headerMap.set("Destination", i);
    }

    if (
      header.includes("valid from") ||
      header.includes("start date") ||
      header.includes("effective date")
    ) {
      headerMap.set("Valid from", i);
    }

    if (
      header.includes("valid to") ||
      header.includes("end date") ||
      header.includes("expiry date")
    ) {
      headerMap.set("Valid to", i);
    }

    if (
      header.includes("20st") ||
      header.includes("20") ||
      header.includes("20 standard")
    ) {
      headerMap.set("20ST", i);
    }

    if (header.includes("40st") || header.includes("40 standard")) {
      headerMap.set("40ST", i);
    }

    if (header.includes("40hc") || header.includes("40 high container")) {
      headerMap.set("40HC", i);
    }

    if (
      header.includes("45hc") ||
      header.includes("45") ||
      header.includes("45 high container")
    ) {
      headerMap.set("45HC", i);
    }

    if (header.includes("soc") || header.includes("shipper owned container")) {
      // Second normalize the header to remove any extra spaces
      headerMap.set("SOC", i);
    }

    // Third check for hazardous
    if (
      header.includes("haz") ||
      header.includes("hazardous") ||
      header.includes("dg")
    ) {
      headerMap.set("HAZ", i);
    }

    // Last check for NOR
    if (
      header.includes("nor") ||
      header.includes("non-reff") ||
      header.includes("nref")
    ) {
      headerMap.set("NOR", i);
    }
  }
  // Log if any of the headers are not found
  const requiredKeys = [
    "SOC",
    "HAZ",
    "NOR",
    "Load",
    "Discharge",
    "Destination",
    "Valid from",
    "Valid to",
    "20ST",
    "40ST",
    "40HC",
    "45HC",
  ];
  for (const key of requiredKeys) {
    if (!headerMap.has(key)) {
      await sendErrorMessage(
        `ERROR IN QHOF Parser, Missing header key: ${key}`
      );
    }
  }

  // If no headers found for Load
  if (!headerMap.has("Load")) {
    headerMap.set("Load", 2);
  }

  // If no headers found for Discharge
  if (!headerMap.has("Discharge")) {
    headerMap.set("Discharge", 3);
  }

  // If no headers found for Discharge
  if (!headerMap.has("Destination")) {
    headerMap.set("Discharge", 4);
  }

  // If no headers found for SOC, use the default header row
  if (!headerMap.has("SOC")) {
    headerMap.set("SOC", 5);
  }

  // If no headers found for NOR, use the default header row
  if (!headerMap.has("NOR")) {
    headerMap.set("NOR", 6);
  }

  // If no headers found for HAZ, use the default header row
  if (!headerMap.has("HAZ")) {
    headerMap.set("HAZ", 7);
  }

  // If no headers found for valid from, use the default header row
  if (!headerMap.has("Valid from")) {
    headerMap.set("Valid from", 8);
  }

  // If no headers found for Valid to, use the default header row
  if (!headerMap.has("Valid to")) {
    headerMap.set("Valid to", 9);
  }

  // If no headers found for 20ST, use the default header row
  if (!headerMap.has("20ST")) {
    headerMap.set("20ST", 11);
  }

  // If no headers found for 40ST, use the default header row
  if (!headerMap.has("40ST")) {
    headerMap.set("40ST", 12);
  }

  // If no headers found for 40HC, use the default header row
  if (!headerMap.has("40HC")) {
    headerMap.set("40HC", 13);
  }

  // If no headers found for 45HC, use the default header row
  if (!headerMap.has("45HC")) {
    headerMap.set("45HC", 14);
  }

  return headerMap;
}

export async function parseQHOFFile(workbook) {
  await initializePorts();
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  // --- Find carrier_contract_number dynamically ---
  let contractNumber = "QHOF-Contract"; // fallback default
  // Scan the first 8 rows to find the carrier contract number
  for (let i = 0; i < 8; i++) {
    const row = jsonData[i];
    if (!row) continue;
    for (const cell of row) {
      if (typeof cell === "string") {
        const match = cell.match(/\bQHOF\w*/);
        if (match) {
          contractNumber = match[0];
          break;
        }
      }
    }
  }

  // --- Find the column index for the header ---
  const headerMap = await getHeaderMap(workbook);

  let lastRow = 0;

  // Process rows sequentially to handle async operations
  let cleanedData = [];
  for (let rowIndex = 0; rowIndex < jsonData.slice(6).length; rowIndex++) {
    const row = jsonData.slice(6)[rowIndex];
    if (row.length === 0) continue;
    lastRow = rowIndex + 6;

    const rowData = {
      row_number: rowIndex + 7,
      rate_source: 653309,
      carrier: 653309,
      transit_time: "",
      contract: "33",
      carrier_contract_number: contractNumber,
      service: "",
    };

    // Process each column
    for (let colIndex = 0; colIndex < row.length; colIndex++) {
      const value = row[colIndex];
      switch (colIndex) {
        case headerMap.get("Load"):
          rowData["port_origin"] = value;
          rowData["port_origin_name"] = value.split(",")[0].trim();
          break;
        case headerMap.get("Discharge"):
          rowData["port_discharge"] = await getPortId(value);
          rowData["port_discharge_name"] = value.split(",")[0].trim();
          break;
        case headerMap.get("Destination"):
          rowData["port_destination"] = await getPortId(
            value ? value.split("\n")[0].trim() : value
          );
          rowData["port_destination_name"] = value
            ? value.split("\n")[0].trim()
            : value;
          break;
        case headerMap.get("Valid from"):
          rowData["vf"] = excelDateToJSDate(value);
          break;
        case headerMap.get("Valid to"):
          rowData["vt"] = excelDateToJSDate(value);
          break;
        case headerMap.get("20ST"):
          rowData["20p"] = value;
          break;
        case headerMap.get("40ST"):
          rowData["40p"] = value;
          break;
        case headerMap.get("40HC"):
          rowData["40hpq"] = value;
          break;
        case headerMap.get("45HC"):
          rowData["45HC"] = value;
          break;
        case headerMap.get("SOC"):
          rowData["SOC"] = value;
          break;
        case headerMap.get("NOR"):
          rowData["NOR"] = value;
          break;
        case headerMap.get("HAZ"):
          rowData["HAZ"] = value;
          break;
      }
    }

    // If port_destination is not provided, use port_discharge
    if (!rowData.port_destination && !rowData.port_destination_name) {
      rowData.port_destination = rowData.port_discharge;
    }

    cleanedData.push(rowData);
  }

  // Filter out empty objects
  cleanedData = cleanedData.filter((row) => Object.keys(row).length > 0);

  // Process port_origin splits
  let processedData = [];
  for (const row of cleanedData) {
    if (row.port_origin && row.port_origin.includes("\n")) {
      const ports = row.port_origin.split("\n");

      for (const port of ports) {
        processedData.push({
          ...row,
          port_origin: await getPortId(port),
          port_origin_name: port.split(",")[0].trim(),
        });
      }
    } else if (row.port_origin) {
      processedData.push({
        ...row,
        port_origin: await getPortId(row.port_origin),
        port_origin_name: row.port_origin_name,
      });
    }
  }

  // Add in the vessel service names
  // We pass the whole processedData to avoid reiterating afterwards
  // This function will populate the service field in each object
  await getVesselServiceName(processedData);

  // Add in the container maintenance charges
  const containerMaintenanceCharges =
    await parseContainerMaintenanceCharges(workbook);
  for (const row of processedData) {
    const matchingCharge = containerMaintenanceCharges.find(
      (charge) =>
        charge.loadPortId === row.port_origin &&
        charge.dischargePortId === row.port_discharge &&
        charge.deliveryPortId === row.port_destination &&
        (charge.original_price_20st !== undefined &&
        charge.original_price_20st !== 0
          ? charge.original_price_20st === row["20p"]
          : true) &&
        (charge.original_price_40st !== undefined &&
        charge.original_price_40st !== 0
          ? charge.original_price_40st === row["40p"]
          : true) &&
        (charge.original_price_40hc !== undefined &&
        charge.original_price_40hc !== 0
          ? charge.original_price_40hc === row["40hpq"]
          : true) &&
        (charge.original_price_45hc !== undefined &&
        charge.original_price_45hc !== 0
          ? charge.original_price_45hc === row["45HC"]
          : true)
    );
    if (matchingCharge) {
      // Add container maintenance charges if present
      if (typeof matchingCharge["20p"] === "number")
        row["20p"] += matchingCharge["20p"];
      if (typeof matchingCharge["40p"] === "number")
        row["40p"] += matchingCharge["40p"];
      if (typeof matchingCharge["40hpq"] === "number")
        row["40hpq"] += matchingCharge["40hpq"];
      if (typeof matchingCharge["45HC"] === "number")
        row["45HC"] += matchingCharge["45HC"];
    }
  }

  processedData = processedData.filter(
    (entry) =>
      (entry.SOC === "" || entry.SOC == null) &&
      (entry.NOR === "" || entry.NOR == null) &&
      (entry.HAZ === "" || entry.HAZ == null)
  );

  // Filter out entries with empty or null port_origin, port_discharge, or port_destination
  processedData = processedData.filter(
    (entry) =>
      entry["port_origin"] !== "" &&
      entry["port_origin"] !== null &&
      entry["port_discharge"] !== "" &&
      entry["port_discharge"] !== null &&
      entry["port_destination"] !== "" &&
      entry["port_destination"] !== null
  );

  deduplicateArray(processedData);

  // Get a set of unique port_origin IDs of the main routes
  const uniqueOriginPortIds = new Set(
    processedData.map((entry) => entry.port_origin)
  );

  // Parse and add feeder rates
  const feederTariffs = await parseFeederRates(workbook);

  let uniqueFeederTariffs = [];
  // Output the feeder rates to a JSON file
  for (const feederRate of feederTariffs) {
    if (uniqueOriginPortIds.has(feederRate.port_origin_id)) continue;
    uniqueFeederTariffs.push(feederRate);
  }

  // Write the feeder rates to a JSON file
  // writeFileSync(
  //   join(dirname(fileURLToPath(import.meta.url));, "../output/unique_feeder_rates.json"),
  //   JSON.stringify(uniqueFeederTariffs, null, 2),
  //   "utf8"
  // );

  const matches = findFeederMatches(processedData, feederTariffs);

  // Combine the new routes (matches) with the existing processedData
  // The new routes are those that have been matched with feeder rates
  // This final array will contain all the main routes (from processedData)
  // and the feeder routes (from matches)

  const combinedData = [...processedData, ...matches];

  // Deduplicate again
  deduplicateArray(combinedData);

  // writeJSONToExcel(combinedData, feederTariffs, processedData);

  // tester(combinedData, processedData, uniqueFeederTariffs);

  fillMissingValues(combinedData);

  const formattedData = formatParsedData(combinedData);

  return formattedData;
}

/**
 *
 * @param {Array of JSON} parsedData data that is parsed and populated from `parseExcelFile`
 * @returns Final formatted data according to JSON that Make is expecting
 */
function formatParsedData(parsedData) {
  return parsedData.map((row) => ({
    carrier: String(row.carrier ?? ""),
    port_destination: String(row.port_destination ?? ""),
    port_origin: String(row.port_origin ?? ""),
    port_discharge: String(row.port_discharge ?? ""),
    vf: String(row.vf ?? ""),
    vt: String(row.vt ?? ""),
    transit_time: String(row.transit_time ?? ""),
    rate_source: String(row.rate_source ?? ""),
    "40p": String(row["40p"] ?? ""),
    "40hqp": String(row["40hpq"] ?? ""),
    "20p": String(row["20p"] ?? ""),
    "45HC": String(row["45HC"] ?? ""),
    service: String(row.service ?? ""),
    contract: String(row.contract ?? ""),
    carrier_contract_number: String(row.carrier_contract_number ?? ""),
  }));
}
