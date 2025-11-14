import {
  readFileSync,
  writeFileSync,
  promises as fsPromises,
  existsSync,
} from "fs";
import axios from "axios";
import * as XLSX from "xlsx";
import { sendErrorMessage } from "./resources/api.service.js";
import { Actor } from "apify";

/**
 * Function to get the latest rate sheet from SFTP
 * @returns Filename / null
 */
async function getFile(ratesheetUrl) {
  try {
    const result = await axios.request({
      responseType: "arraybuffer",
      url: ratesheetUrl || "https://www.primefreight.com/cma_rates/ratesheet.xlsx",
      method: "get",
      headers: {
        "Content-Type": "blob",
      },
    });
    const outputFilename = "ratesheet.xlsx";
    writeFileSync(outputFilename, result.data);
    console.log(`File downloaded and saved as: ${outputFilename}`);
    return outputFilename;
  } catch (error) {
    await sendErrorMessage(`ERROR: Unable to obtain file from SFTP: ${error}`);
    console.error(`ERROR: Failed to download file: ${error}`);
    return null;
  }
}

function isSVContract(workbook, fileName) {
  // Look for a sheet containing "cover" (case-insensitive)
  const sheetName = workbook.SheetNames.find(name => name.toLowerCase().includes("cover"));
  if (!sheetName) {
    console.log(`File ${fileName}: No sheet with 'cover' in name found for SVC check`);
    console.log(`Available sheets: ${workbook.SheetNames.join(", ")}`);
    return false;
  }
  const coverSheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(coverSheet, { header: 1, defval: "" });

  console.log(`File ${fileName}: Checking '${sheetName}' sheet for Service Contract #: pattern`);
  // Search for "Service Contract" and check for 3117 or 3118
  for (const row of jsonData) {
    if (!row) continue;
    const rowString = row.map(cell => (cell !== null && cell !== undefined ? cell.toString().toLowerCase() : ""));
    if (rowString.some(cell => cell.includes("service contract"))) {
      console.log(`File ${fileName}: Found 'Service Contract' in row: ${JSON.stringify(row)}`);
      // Check entire sheet for 3117 or 3118 to be safe
      for (const innerRow of jsonData) {
        if (!innerRow) continue;
        for (const cell of innerRow) {
          if (cell === null || cell === undefined) continue;
          const cellValue = cell.toString();
          if (cellValue.includes("3117") || cellValue.includes("3118")) {
            console.log(`File ${fileName}: Matches SVC criteria (found ${cellValue} in Cover sheet)`);
            return true;
          }
        }
      }
    }
  }
  console.log(`File ${fileName}: Does not match SVC criteria (3117 or 3118 not found in Cover sheet)`);
  return false;
}

function isQHOFContract(workbook, fileName) {
  const sheetName = workbook.SheetNames[0];
  const firstSheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

  // Start reading the file
  for (const row of jsonData) {
    if (!row) continue;

    // Read through each cell
    for (const cell of row) {
      if (typeof cell === "string") {
        // If the cell matches QHOF
        const match = cell.match(/\bQHOF\w*/);
        if (match) {
          console.log(`File ${fileName}: Matches QHOF criteria (QHOF pattern found in content)`);
          return true;
        }
      }
    }
  }
  console.log(`File ${fileName}: Does not match QHOF criteria (no QHOF pattern found in content)`);
  return false;
}

async function forwardSVContract({ workbook, fileName, mode }) {
  try {
    // Decide runner by explicit mode or detect via cover sheet
    let target = mode;
    if (!target) {
      const sheetName = workbook.SheetNames.find((name) =>
        name.toLowerCase().includes("cover")
      );
      let is3117 = false;
      if (sheetName) {
        const coverSheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(coverSheet, {
          header: 1,
          defval: "",
        });
        for (const row of jsonData) {
          if (!row) continue;
          for (const cell of row) {
            if (cell === null || cell === undefined) continue;
            const cellValue = cell.toString();
            if (cellValue.includes("3117")) {
              is3117 = true;
              break;
            }
          }
          if (is3117) break;
        }
      }
      target = is3117 ? "SVC_3117_CONTRACT" : "SVC_3117_FEEDER";
    }

    if (target === "SVC_3117_CONTRACT") {
      await import("./SVC-3117-contract/main.js");
      console.log(`Forwarded to SVC-3117-contract/main.js for file: ${fileName}`);
    } else if (target === "SVC_3117_FEEDER") {
      await import("./SVC-3117-Feeder/main.js");
      console.log(`Forwarded to SVC-3117-Feeder/main.js for file: ${fileName}`);
    } else {
      throw new Error(`Unknown mode: ${target}`);
    }
  } catch (error) {
    await sendErrorMessage(`ERROR in forwarding SVC contract for file ${fileName}: ${error}`);
    console.error(`ERROR: Failed to forward SVC contract for file ${fileName}: ${error}`);
  }
}

async function main() {
  const input = (await Actor.getInput()) || {};
  const {
    ratesheetUrl,
    mode, // "AUTO" | "SVC_3117_CONTRACT" | "SVC_3117_FEEDER"
    delayBeforeMs = 0,
    pushResults = true,
  } = input;

  if (delayBeforeMs && Number(delayBeforeMs) > 0) {
    console.log(`Waiting ${delayBeforeMs} ms before processing...`);
    await new Promise((res) => setTimeout(res, Number(delayBeforeMs)));
  }
  // 1. Get the file
  const fileName = await getFile(ratesheetUrl);

  if (fileName === null) {
    await sendErrorMessage(`Filename from SFTP is null`);
    console.error(`ERROR: No file to process, filename is null`);
    return;
  }

  const fileData = readFileSync(`./${fileName}`);
  console.log(`Reading file: ${fileName}`);

  const workbook = XLSX.read(fileData, { type: "buffer" });
  // At this point, the file is open
  console.log(`File ${fileName}: Successfully loaded into workbook`);

  if (isSVContract(workbook, fileName)) {
    console.log(`SVC Contract identified for file: ${fileName}`);
    await forwardSVContract({ workbook, fileName, mode: mode === "AUTO" ? undefined : mode });
  } else {
    // Send error message and log if neither QHOF nor SVC
    const errorMsg = `Contract was not identified and therefore not sent out for file: ${fileName}`;
    await sendErrorMessage(errorMsg);
    console.log(errorMsg);
  }

  try {
    await fsPromises.unlink(`./${fileName}`);
    console.log(`File ${fileName}: Successfully deleted`);
  } catch (error) {
    await sendErrorMessage(`APIFY: Error deleting file ${fileName}: ${error}`);
    console.error(`ERROR: Failed to delete file ${fileName}: ${error}`);
  }

  if (pushResults) {
    const candidates = [
      "SVC-3117-contract/FinalRatesToEndpoint.json",
      "SVC-3117-Feeder/FinalRatesToEndpoint.json",
      "FinalRatesToEndpoint.json",
    ];
    for (const p of candidates) {
      try {
        if (existsSync(p)) {
          const json = JSON.parse(readFileSync(p, "utf8"));
          if (Array.isArray(json)) {
            await Actor.pushData(json);
            console.log(`Pushed ${json.length} items from ${p} to dataset.`);
          } else {
            await Actor.pushData(json);
            console.log(`Pushed object result from ${p} to dataset.`);
          }
          break;
        }
      } catch (e) {
        console.warn(`Could not push results from ${p}: ${e}`);
      }
    }
  }
}

// Actor initialization
await Actor.init();

console.log("Hello from actor");

await main().catch(async (error) => {
  const errorMessage = `Error thrown running main: ${error}`;
  console.error(errorMessage);
  await sendErrorMessage(errorMessage);
});

await Actor.exit();
