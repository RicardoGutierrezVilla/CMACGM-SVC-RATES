import XLSX from "xlsx";
import { getPortId, normalize } from "../resources/utils.js";

export async function parseContainerMaintenanceCharges(workbook) {
  const sheet = workbook.Sheets["Standard charges"];
  if (!sheet) {
    console.error("Sheet 'Standard charges' not found.");
    return;
  }
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

  // First find the header row:
  const headerRowIndex = data.findIndex((row) =>
    row.some((cell) =>
      // Look for the presence of any of the following headers in the whole spreadsheet
      // This indicates the header row for container maintenance
      [
        "place of receipt",
        "load port",
        "discharge port",
        "place of delivery",
      ].some((header) => normalize(cell).includes(normalize(header)))
    )
  );

  if (headerRowIndex === -1) {
    console.error("Header row not found for container maintenance.");
    return;
  }

  // 1. Identify column indexes for headers
  const headerRow = data[headerRowIndex];
  const headersToFind = ["20ST", "40ST", "40HC", "45HC", "Charge Description"];
  const headerCols = {};
  headersToFind.forEach((header) => {
    const colIdx = headerRow.findIndex(
      (cell) =>
        cell && cell.toString().trim().toLowerCase() === header.toLowerCase()
    );
    if (colIdx !== -1) headerCols[header] = colIdx;
  });

  // Get the headers for the port
  const portHeaders = ["Load Port", "Discharge Port", "Place Of\nDelivery"];
  const portCols = {};
  portHeaders.forEach((header) => {
    const colIdx = headerRow.findIndex(
      (cell) =>
        cell && cell.toString().trim().toLowerCase() === header.toLowerCase()
    );
    if (colIdx !== -1) portCols[header] = colIdx;
  });

  // Helper to get merged cell value for a given row/col
  function getMergedCellValue(row, col) {
    if (!sheet["!merges"]) return data[row][col];
    for (const merge of sheet["!merges"]) {
      if (
        row >= merge.s.r &&
        row <= merge.e.r &&
        col >= merge.s.c &&
        col <= merge.e.c
      ) {
        const cellRef = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
        return sheet[cellRef] ? sheet[cellRef].v : undefined;
      }
    }
    return data[row][col];
  }

  // Find all block start rows by looking for changes in the "Load Port" merged cell value
  // Used to get each entry based on the merged cell rather than a prefixed values of rows
  let lastLoadPort = null;
  let lastDischargePort = null;
  let lastDelivery = null;
  const blockStarts = [];
  // for (let row = 2; row < data.length; row++) {
  //   const loadPortVal = getMergedCellValue(row, portCols["Load Port"]);
  //   const dischargePortVal = getMergedCellValue(
  //     row,
  //     portCols["Discharge Port"]
  //   );
  //   const deliveryVal = getMergedCellValue(row, portCols["Place Of\nDelivery"]);
  //   if (
  //     (loadPortVal && loadPortVal !== lastLoadPort) ||
  //     (dischargePortVal && dischargePortVal !== lastDischargePort) ||
  //     (deliveryVal && deliveryVal !== lastDelivery)
  //   ) {
  //     blockStarts.push(row);
  //     lastLoadPort = loadPortVal;
  //     lastDischargePort = dischargePortVal;
  //     lastDelivery = deliveryVal;
  //   }
  // }

  for (let row = 0; row < data.length; row++) {
    for (let col = 0; col < data[row].length; col++) {
      const cell = data[row][col];
      if (
        cell &&
        typeof cell === "string" &&
        cell.trim().toLowerCase().includes("container charges")
      ) {
        // If we find a row with "container charges", we need to add it to the block starts
        blockStarts.push(row);
        break;
      }
    }
  }

  // After populating blockStarts
  if (
    blockStarts.length === 0 ||
    blockStarts[blockStarts.length - 1] < data.length - 1
  ) {
    // Find the first non-empty row after the last detected block
    for (
      let row = (blockStarts[blockStarts.length - 1] || 0) + 1;
      row < data.length;
      row++
    ) {
      if (
        data[row] &&
        data[row].some((cell) => cell && cell.toString().trim() !== "")
      ) {
        blockStarts.push(row);
        break;
      }
    }
  }
  // Add the last block if it exists
  const results = [];
  for (let b = 0; b < blockStarts.length; b++) {
    const blockStart = blockStarts[b];
    const blockEnd =
      b + 1 < blockStarts.length ? blockStarts[b + 1] : data.length;

    const loadPortVal = getMergedCellValue(blockStart, portCols["Load Port"]);
    const dischargePortVal = getMergedCellValue(
      blockStart,
      portCols["Discharge Port"]
    );
    const deliveryVal = getMergedCellValue(
      blockStart,
      portCols["Place Of\nDelivery"]
    );

    // If the load port is empty, skip this block
    const getPortName = (val, isDelivery = false) => {
      if (!val) return "";
      const lines = val
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean);
      if (isDelivery) return lines[0];
      return lines[lines.length - 1];
    };

    // Clean the port names
    const load_port = getPortName(loadPortVal);
    const discharge_port = getPortName(dischargePortVal);
    const delivery = getPortName(deliveryVal, true);

    // Get the port id's
    const loadPortId = await getPortId(load_port);
    const dischargePortId = await getPortId(discharge_port);
    let deliveryPortId = await getPortId(delivery);

    // If the delivery port is empty, use the discharge port id
    if (delivery === "" && deliveryPortId === null) {
      deliveryPortId = dischargePortId;
    }

    // Build the base object with port info
    const entry = {
      block_number: b + 1,
      load_port,
      discharge_port,
      delivery,
      loadPortId,
      dischargePortId,
      deliveryPortId,
    };

    for (let rowIdx = blockStart; rowIdx < blockEnd; rowIdx++) {
      const row = data[rowIdx];
      if (!row) continue;
      const chargeDescCol = headerCols["Charge Description"];
      if (chargeDescCol === undefined) continue;
      const chargeDesc = row[chargeDescCol]?.toString().toLowerCase();

      // Only process rows with a relevant charge description
      if (
        chargeDesc &&
        (chargeDesc.includes("container maintenance charge") ||
          chargeDesc.includes("rate offer per container"))
      ) {
        // Add container maintenance charge fields if present
        if (chargeDesc.includes("container maintenance charge")) {
          entry["20p"] =
            row[headerCols["20ST"]] !== undefined &&
            row[headerCols["20ST"]] !== ""
              ? parseFloat(row[headerCols["20ST"]].replace(/,/g, ""))
              : 0;
          entry["40p"] =
            row[headerCols["40ST"]] !== undefined &&
            row[headerCols["40ST"]] !== ""
              ? parseFloat(row[headerCols["40ST"]].replace(/,/g, ""))
              : 0;
          entry["40hpq"] =
            row[headerCols["40HC"]] !== undefined &&
            row[headerCols["40HC"]] !== ""
              ? parseFloat(row[headerCols["40HC"]].replace(/,/g, ""))
              : 0;
          entry["45HC"] =
            row[headerCols["45HC"]] !== undefined &&
            row[headerCols["45HC"]] !== ""
              ? parseFloat(row[headerCols["45HC"]].replace(/,/g, ""))
              : 0;
        }

        // Add original price fields if present
        if (
          normalize(chargeDesc).includes("rate offer per container") ||
          normalize(chargeDesc).includes("rate offer")
        ) {
          entry["original_price_20st"] =
            row[headerCols["20ST"]] !== undefined &&
            row[headerCols["20ST"]] !== ""
              ? parseFloat(row[headerCols["20ST"]].replace(/,/g, ""))
              : 0;
          entry["original_price_40st"] =
            row[headerCols["40ST"]] !== undefined &&
            row[headerCols["40ST"]] !== ""
              ? parseFloat(row[headerCols["40ST"]].replace(/,/g, ""))
              : 0;
          entry["original_price_40hc"] =
            row[headerCols["40HC"]] !== undefined &&
            row[headerCols["40HC"]] !== ""
              ? parseFloat(row[headerCols["40HC"]].replace(/,/g, ""))
              : 0;
          entry["original_price_45hc"] =
            row[headerCols["45HC"]] !== undefined &&
            row[headerCols["45HC"]] !== ""
              ? parseFloat(row[headerCols["45HC"]].replace(/,/g, ""))
              : 0;
        }
      }
    }
    // Check to make sure we have at least one of the container maintenance charge fields
    if (
      typeof entry["original_price_20st"] === "number" ||
      typeof entry["original_price_40st"] === "number" ||
      typeof entry["original_price_40hc"] === "number" ||
      typeof entry["original_price_45hc"] === "number"
    ) {
      results.push(entry);
    }
  }

  return results;
}
