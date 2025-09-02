import fs from "fs";
import XLSX from "xlsx";
import path from "path";

export function tester(parsedDataRates, mainRoutes, uniqueFeederRoutes) {
  // Create a lookup for feeder rates by port_origin_id
  const feederLookup = {};
  for (const feeder of uniqueFeederRoutes) {
    const originId = feeder.port_origin_id;
    const polId = feeder.main_pol_id;

    if (!feederLookup[originId]) {
      feederLookup[originId] = {};
    }
    // If you want to keep only the lowest rate per POL, use Math.min as before:
    if (!feederLookup[originId][polId]) {
      feederLookup[originId][polId] = { ...feeder };
    } else {
      // For each container size, keep the lowest value
      const existing = feederLookup[originId][polId];
      if ((feeder["20p"] ?? 0) < (existing["20p"] ?? 0)) {
        existing["20p"] = feeder["20p"];
        existing["main_pol_name"] = feeder["main_pol_name"];
      }
      if ((feeder["40p"] ?? 0) < (existing["40p"] ?? 0)) {
        existing["40p"] = feeder["40p"];
      }
      if ((feeder["40hpq"] ?? 0) < (existing["40hpq"] ?? 0)) {
        existing["40hpq"] = feeder["40hpq"];
      }
      if ((feeder["45HC"] ?? 0) < (existing["45HC"] ?? 0)) {
        existing["45HC"] = feeder["45HC"];
      }
    }
  }

  const results = [];

  for (const entry of parsedDataRates) {
    const feeders = feederLookup[entry.port_origin];
    if (feeders) {
      for (const polId in feeders) {
        const feeder = feeders[polId];
        results.push({
          ...entry,
          "20p_subtracted": entry["20p"] - (feeder["20p"] ?? 0),
          "40p_subtracted": entry["40p"] - (feeder["40p"] ?? 0),
          "40hpq_subtracted": entry["40hpq"] - (feeder["40hpq"] ?? 0),
          "45hc_subtracted": entry["45HC"] - (feeder["45HC"] ?? 0),
          main_pol_id: feeder["main_pol_id"],
          POL: feeder["main_pol_name"],
          matched_port: null,
          matched: null,
        });
      }
    } else {
      // If no feeder match, just copy the entry and set subtracted fields to null
      results.push({
        ...entry,
        "20p_subtracted": null,
        "40p_subtracted": null,
        "40hpq_subtracted": null,
        "45hc_subtracted": null,
        matched: false,
        matched_port: null,
      });
    }
  }

  for (const entry of mainRoutes) {
    for (const rate of results) {
      if (entry.port_origin === rate.main_pol_id) {
        if (
          rate["20p_subtracted"] === entry["20p"] &&
          rate["40p_subtracted"] === entry["40p"] &&
          rate["40hpq_subtracted"] === entry["40hpq"] &&
          rate["45hc_subtracted"] === entry["45HC"]
        ) {
          rate.matched = true;
          rate.matched_port = entry.port_origin_name;
        }
      }
    }
  }

  // Clean up at the end
  for (const entry of results) {
    if (entry.matched === null && entry.matched_port === null) {
    }
  }

  const feederSummaryRows = [];
  for (const originId in feederLookup) {
    const pols = Object.values(feederLookup[originId]);
    if (pols.length > 1) {
      // Find cheapest for each container size
      const cheapest = {
        "20p": pols.reduce(
          (min, curr) => (curr["20p"] < min["20p"] ? curr : min),
          pols[0]
        ),
        "40p": pols.reduce(
          (min, curr) => (curr["40p"] < min["40p"] ? curr : min),
          pols[0]
        ),
        "40hpq": pols.reduce(
          (min, curr) => (curr["40hpq"] < min["40hpq"] ? curr : min),
          pols[0]
        ),
        "45HC": pols.reduce(
          (min, curr) => (curr["45HC"] < min["45HC"] ? curr : min),
          pols[0]
        ),
      };
      for (const pol of pols) {
        feederSummaryRows.push({
          port_origin_id: pol.port_origin_id,
          port_origin_name: pol.port_origin_name,
          main_pol_id: pol.main_pol_id,
          main_pol_name: pol.main_pol_name,
          "20p": pol["20p"],
          "20p_is_cheapest": pol === cheapest["20p"],
          "40p": pol["40p"],
          "40p_is_cheapest": pol === cheapest["40p"],
          "40hpq": pol["40hpq"],
          "40hpq_is_cheapest": pol === cheapest["40hpq"],
          "45HC": pol["45HC"],
          "45HC_is_cheapest": pol === cheapest["45HC"],
        });
      }
    }
  }

  const feederSummary = XLSX.utils.json_to_sheet(feederSummaryRows);

  // Write to Excel
  const worksheet = XLSX.utils.json_to_sheet(results);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "SubtractedRates");
  XLSX.utils.book_append_sheet(workbook, feederSummary, "Summmary Sheet");

  const outputPath = path.resolve("../output/qho_rates.xlsx");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  fs.writeFileSync(outputPath, buffer);

  return results;
}
