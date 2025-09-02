import XLSX from "xlsx";
import { getPortId, initializePorts, normalize } from "../resources/utils.js";
import {
  writeFileSync,
  existsSync,
  readFileSync,
  appendFileSync,
  mkdirSync,
} from "fs";
import { initializeSCACCodes } from "../resources/api.service.js";
import { dirname } from "path";

const logFilePath = "../output/feeder_port_not_found.log";
export const genericWords = ["port", "tanjung", "st"];
let scasCodes = null;
let scacCodesPromise = null;

async function getSCACCodes() {
  // Check if SCAC codes are already initialized
  if (scasCodes) {
    return scasCodes;
  }
  if (!scacCodesPromise) {
    // Initialize SCAC codes if not already done
    scacCodesPromise = initializeSCACCodes();
  }
  scasCodes = await scacCodesPromise;

  return scasCodes;
}

async function findPortId(portName, rowNumber) {
  if (!portName) return null;

  const portsCache = await initializePorts();

  if (genericWords.some((word) => portName.includes(word))) {
    return null;
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

async function findAndMatchPort(portName, rowNumber) {
  // If the port name is empty, return null
  if (!portName) return null;

  await getSCACCodes();

  // Find the SCAC code before anything
  // Usually SCAS codes are the top line of the port name
  const parts = portName
    .split(/\s*[\r\n]+\s*/)
    .map((p) => p.trim())
    .filter(Boolean);

  if (!parts[0]) {
    const notFoundMsg = `${rowNumber}: No SCAC code found in portName: "${portName.replace(
      /\n/g,
      " "
    )}"\n`;
    console.error(notFoundMsg);
    return null;
  }

  const scasCode = parts[0];
  const portIdByScacCode = scasCodes[scasCode.toUpperCase()];

  if (portIdByScacCode) {
    const foundMsg = `${rowNumber}: Port ID found by SCAC code: ${portIdByScacCode}, Port Name: ${portName.replace(
      /\n/g,
      " "
    )}, Parsed SCAC code: ${scasCode}\n`;
    return portIdByScacCode;
  } else {
    const notFoundMsg = `${rowNumber}: SCAC code not found: ${scasCode}, Port Name: ${portName.replace(
      /\n/g,
      " "
    )}\n`;
  }

  const ports = await initializePorts();
  // First split on the new line character
  // Format for this section is usually
  // "Code for port name\nPort name"
  const portNames = portName
    // split on newline, comma, or space (one or more)
    .split(/[\n, ]+/)
    .map((p) => p.trim())
    .filter(Boolean)
    .filter(
      (name) => !genericWords.some((word) => normalize(name).includes(word))
    );

  // If nothing matched, try to find the port name based on each word
  // Again with a preference for the latter words as they usually contain the city
  portName = normalize(portName);
  const words = portName.split(" ").filter(Boolean);

  // First, try to match the port name directly, after splitting by new line
  words.forEach((word, index) => {
    if (word.includes("\n")) {
      words[index] = word.split("\n").filter(Boolean);
    }
  });

  // Try each word, prefer the last word (usually the city)
  for (let i = words.length - 1; i >= 0; i--) {
    const word = words[i];
    if (word) {
      // If the word is a generic word, skip it
      if (
        genericWords.forEach((genericWord) =>
          normalize(word).includes(genericWord)
        )
      )
        continue;
      const portId = await findPortId(word, rowNumber);
      if (portId) {
        return portId;
      }
    }
  }

  // If nothing matched, try the full portName
  const fallbackPortId = await getPortId(portName);
  if (fallbackPortId) {
    return fallbackPortId;
  }

  return null;
}

function mapContainerType(containerType, rowNumber) {
  if (!containerType) return null;
  switch (containerType.trim().toUpperCase()) {
    case "20ST":
      return "20p";
    case "40ST":
      return "40p";
    case "40HC":
      return "40hpq";
    case "45HC":
      return "45HC";
    default:
      return containerType;
  }
}

export async function parseFeederRates(workbook) {
  let sheet = workbook.Sheets["Feeder tariff book"];

  if (!sheet) {
    // Try to find another sheet containing "feeder"
    const feederSheetName = Object.keys(workbook.Sheets).find((name) =>
      normalize(name).includes("feeder")
    );
    if (feederSheetName) {
      sheet = workbook.Sheets[feederSheetName];
    } else {
      console.error("No sheet found containing 'feeder'");
      return;
    }
  }

  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  // 1. Find the header row
  // Assuming the header row contains one of the following strings:
  // "Feeder tariff", "out port", "main pol", "main pod"
  const headerRowIndex = data.findIndex((row) =>
    row.some((cell) =>
      ["out port", "main pol", "main pod"].some((header) =>
        normalize(cell).includes(normalize(header))
      )
    )
  );

  // 2. Identify column indexes for headers -- this will be used later on in the program to collect data
  // -- To make sure that the data is still parsed regardless of the order of the columns
  const headerRow = data[headerRowIndex];
  const headersToFind = ["out port", "main pol", "main pod", "eq", "rate"];
  const headerCols = {};
  headersToFind.forEach((header) => {
    const colIdx = headerRow.findIndex(
      (cell) => cell && normalize(cell).includes(normalize(header))
    );
    if (colIdx !== -1) headerCols[header] = colIdx;
  });

  // 3. Collect data from the rows below the header row
  const rates = [];
  for (let row = headerRowIndex + 1; row < data.length; row++) {
    const rowData = data[row];
    if (rowData.length === 0) continue; // Skip empty rows

    // Initialize the rowData object with default values
    const feederRate = {};
    feederRate["row_number"] = row + 1; // Store the row number for reference
    // Process each column based on the identified header columns
    for (let colIndex = 0; colIndex < rowData.length; colIndex++) {
      switch (colIndex) {
        case headerCols["Feeder tariff"]:
          feederRate["feeder_tariff"] = rowData[colIndex];
          break;
        case headerCols["out port"]:
          feederRate["port_origin_id"] = await findAndMatchPort(
            rowData[colIndex],
            row + 1
          );
          feederRate["port_origin_name"] = rowData[colIndex];
          break;
        case headerCols["main pol"]:
          feederRate["main_pol_id"] = await findAndMatchPort(
            rowData[colIndex],
            row + 1
          );
          feederRate["main_pol_name"] = rowData[colIndex];
          break;
        case headerCols["main pod"]:
          feederRate["main_pod_id"] = await findAndMatchPort(
            rowData[colIndex],
            row + 1
          );
          feederRate["main_pod_name"] = rowData[colIndex];
          break;
        case headerCols["eq"]:
          feederRate["eq"] = mapContainerType(rowData[colIndex], row + 1);
          break;
        case headerCols["rate"]:
          feederRate["rate"] = Number(
            String(rowData[colIndex]).replace(/,/g, "").trim()
          );
          break;
        default:
          // Handle any other columns if necessary
          break;
      }
    }
    // Remove any undefined or null values from the feederRate object
    // This is usually "main pod" and similarly "main pod id"

    // If the origin port is null, output it to a log file
    // If nothing is matched, log the port name in a file:

    /*
      For testing purposes
    */
    // if (feederRate.port_origin_id === null) {
    //   try {
    //     // Ensure the output directory exists
    //     const logDir = dirname(logFilePath);
    //     if (!existsSync(logDir)) {
    //       mkdirSync(logDir, { recursive: true });
    //     }
    //     if (!existsSync(logFilePath)) {
    //       writeFileSync(logFilePath, "");
    //     }
    //     const logContent = readFileSync(logFilePath, "utf8");
    //     const alreadyLogged = logContent
    //       .split("\n")
    //       .some((line) => line.trim() === feederRate.port_origin_name.trim());
    //     if (!alreadyLogged) {
    //       appendFileSync(
    //         logFilePath,
    //         feederRate.port_origin_name.split("\n")[1].trim() + "\n"
    //       );
    //     }
    //   } catch (err) {
    //     console.error("Error logging unfound port:", err);
    //   }
    // }

    if (
      feederRate.main_pod_id === null &&
      feederRate.main_pod_name.trim() === ""
    ) {
      delete feederRate.main_pod_id;
      delete feederRate.main_pod_name;
    }
    if (feederRate.port_origin_id === null) {
      // skip if port_origin_id is null
      continue;
    }

    // 4. Combine rates with the same port_origin_id and main_pol_id
    // Check if the route already exists in the rates array
    // If it already exists, take the higher rate
    // If it doesn't exist, add the new rate
    const existingRate = rates.find(
      (rate) =>
        rate.port_origin_id === feederRate.port_origin_id &&
        rate.main_pol_id === feederRate.main_pol_id
    );
    if (existingRate) {
      // If it exists, update the rate for the corresponding eq

      // Check if the existing rate for the container type is higher
      // 1. Check if the existing rate has the same container type key
      if (
        feederRate.eq &&
        Object.prototype.hasOwnProperty.call(existingRate, feederRate.eq)
      ) {
        // 2. If it does, check if the existing rate is higher
        if (existingRate[feederRate.eq] < feederRate.rate) {
          // 3. If it is, update the existing rate
          existingRate.row_number = feederRate.row_number;
          existingRate[feederRate.eq] = feederRate.rate;
        }
        // 4. If it isn't, do not add the lower price
      }
      // 5. If it doesn't have the same container type key, add the new rate
      else if (feederRate.eq && feederRate.rate !== undefined) {
        existingRate[feederRate.eq] = feederRate.rate;
      }
    }
    // If it doesn't exist, add the new rate
    else {
      // Add the feederRate to the rates array
      // Reorganize the feederRate object
      if (feederRate.eq && feederRate.rate !== undefined) {
        // If eq and rate are defined, add them to the rates array
        feederRate[feederRate.eq] = feederRate.rate;
        delete feederRate.eq;
        delete feederRate.rate;
        rates.push(feederRate);
      }
    }
  }

  // 5. Clean up the rates array
  // Extraneous fields
  for (const rate of rates) {
    if (rate.eq && rate.rate !== undefined) {
      rate[rate.eq] = rate.rate;
    }
    delete rate.eq;
    delete rate.rate;
  }

  return rates;
}
