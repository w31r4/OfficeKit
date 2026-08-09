import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import JSZip from "jszip";

import { SpreadsheetFile, Workbook } from "office-kit";

export function buildPivotTableWorkbook() {
  const workbook = Workbook.create();
  const data = workbook.worksheets.add("Data");
  data.getRange("A1:F13").write([
    ["Region", "Channel", "Product", "Revenue", "Units", "Order date"],
    ["East", "Direct", "Alpha", 70, 7, new Date("2026-07-01T00:00:00Z")],
    ["East", "Direct", "Beta", 50, 5, new Date("2026-07-01T00:00:00Z")],
    ["East", "Partner", "Alpha", 45, 4, new Date("2026-07-15T00:00:00Z")],
    ["East", "Partner", "Beta", 35, 4, new Date("2026-07-15T00:00:00Z")],
    ["West", "Direct", "Alpha", 90, 9, new Date("2026-07-01T00:00:00Z")],
    ["West", "Direct", "Beta", 60, 6, new Date("2026-07-01T00:00:00Z")],
    ["West", "Partner", "Alpha", 55, 5, new Date("2026-07-15T00:00:00Z")],
    ["West", "Partner", "Beta", 35, 4, new Date("2026-07-15T00:00:00Z")],
    ["North", "Direct", "Alpha", 60, 6, new Date("2026-08-01T00:00:00Z")],
    ["North", "Direct", "Beta", 50, 5, new Date("2026-08-01T00:00:00Z")],
    ["North", "Partner", "Alpha", 40, 4, new Date("2026-08-01T00:00:00Z")],
    ["North", "Partner", "Beta", 30, 3, new Date("2026-08-01T00:00:00Z")],
  ]);
  data.getRange("A1:F1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  data.getRange("D2:D13").setNumberFormat("$#,##0");
  data.getRange("E2:E13").setNumberFormat("#,##0");
  data.getRange("F2:F13").setNumberFormat("yyyy-mm-dd");
  data.getRange("A1:F13").format.autofitColumns();
  data.getRange("F1:F13").format.columnWidthPx = 72;
  data.freezePanes.freezeRows(1);
  data.showGridLines = false;

  const summary = workbook.worksheets.add("Pivot Summary");
  summary.getRange("A1:H6").format = { border: { bottom: { style: "thin", color: "#CBD5E1" } } };
  summary.getRange("A1:H1").format = { fill: "#DBEAFE", font: { bold: true, color: "#1E3A8A", size: 9 } };
  summary.getRange("A1:H1").format.wrapText = true;
  summary.getRange("A1:H1").format.rowHeightPx = 42;
  summary.getRange("A6:H6").format = { fill: "#E2E8F0", font: { bold: true, color: "#0F172A" } };
  for (const range of ["C2:C6", "E2:E6", "G2:G6"]) summary.getRange(range).setNumberFormat("$#,##0");
  for (const range of ["D2:D6", "F2:F6", "H2:H6"]) summary.getRange(range).setNumberFormat("#,##0");
  // Keep the eight-column summary within the same 440 px width budget that was
  // stable on GitHub's Linux LibreOffice/metric stack. The smaller header font
  // and explicit wrap height preserve the long value labels without clipping.
  summary.getRange("A1:A6").format.columnWidthPx = 58;
  summary.getRange("B1:B6").format.columnWidthPx = 54;
  summary.getRange("C1:F6").format.columnWidthPx = 46;
  summary.getRange("G1:G6").format.columnWidthPx = 78;
  summary.getRange("H1:H6").format.columnWidthPx = 66;
  summary.showGridLines = false;
  summary.pivotTables.add({
    name: "Revenue and units by region",
    sourceRange: "Data!A1:E13",
    targetRange: "A1",
    rowFields: ["Region", "Channel"],
    columnFields: ["Product"],
    valueFields: [
      { field: "Revenue", summarizeBy: "sum", name: "Revenue" },
      { field: "Units", summarizeBy: "sum", name: "Units" },
    ],
    filters: [{ field: "Region", exclude: ["North"] }],
    rowGrandTotals: true,
    columnGrandTotals: true,
    refreshPolicy: { refreshOnLoad: true, saveData: true, enableRefresh: true },
  });

  const dateSummary = workbook.worksheets.add("Date Summary");
  dateSummary.getRange("A1:B4").format = { border: { bottom: { style: "thin", color: "#CBD5E1" } } };
  dateSummary.getRange("A1:B1").format = { fill: "#DCFCE7", font: { bold: true, color: "#14532D" } };
  dateSummary.getRange("A2:A3").setNumberFormat("yyyy-mm-dd");
  dateSummary.getRange("B2:B4").setNumberFormat("$#,##0");
  dateSummary.getRange("A1:B4").format.autofitColumns();
  dateSummary.showGridLines = false;
  dateSummary.pivotTables.add({
    name: "July revenue by date",
    sourceRange: "Data!A1:F13",
    targetRange: "A1",
    rowFields: ["Order date"],
    valueFields: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue" }],
    filters: [{ field: "Order date", type: "dateBetween", value1: "2026-07-01", value2: "2026-07-31" }],
    columnGrandTotals: true,
    refreshPolicy: { refreshOnLoad: false, saveData: true, enableRefresh: true },
  });
  return workbook;
}

export async function createPivotTableWorkbook(outputPath) {
  const workbook = buildPivotTableWorkbook();
  const summary = workbook.worksheets.getItem("Pivot Summary");
  const pivot = summary.pivotTables.items[0];
  const datePivot = workbook.worksheets.getItem("Date Summary").pivotTables.items[0];
  assert.deepEqual(pivot.computedValues(), [
    ["Region", "Channel", "Alpha — Revenue", "Alpha — Units", "Beta — Revenue", "Beta — Units", "Grand Total — Revenue", "Grand Total — Units"],
    ["East", "Direct", 70, 7, 50, 5, 120, 12],
    ["East", "Partner", 45, 4, 35, 4, 80, 8],
    ["West", "Direct", 90, 9, 60, 6, 150, 15],
    ["West", "Partner", 55, 5, 35, 4, 90, 9],
    ["Grand Total", "", 260, 25, 180, 19, 440, 44],
  ]);
  assert.deepEqual(datePivot.computedValues(), [
    ["Order date", "Revenue"],
    [new Date("2026-07-01T00:00:00Z"), 270],
    [new Date("2026-07-15T00:00:00Z"), 170],
    ["Grand Total", 440],
  ]);

  const inspection = workbook.inspect({ kind: "sheet,pivotTable,style", sheetName: summary.name, range: "A1:H6", maxChars: 16_000 });
  assert.match(inspection.ndjson, /"kind":"pivotTable"/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const preview = await workbook.render({ sheetName: summary.name, range: "A1:H6", format: "svg" });
  assert.match(await preview.text(), /Revenue and units by region/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const firstZip = await JSZip.loadAsync(new Uint8Array(await first.arrayBuffer()));
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedPivot = imported.worksheets.getItem(summary.name).pivotTables.items[0];
  const importedDatePivot = imported.worksheets.getItem("Date Summary").pivotTables.items[0];
  assert.deepEqual(importedPivot.computedValues(), pivot.computedValues());
  assert.deepEqual(importedPivot.filters, [{ field: "Region", exclude: ["North"] }]);
  assert.deepEqual(importedDatePivot.filters, [{ field: "Order date", type: "dateBetween", value1: "2026-07-01", value2: "2026-07-31", useWholeDay: true }]);
  assert.deepEqual(importedDatePivot.computedValues(), [
    ["Order date", "Revenue"],
    [46_204, 270],
    [46_218, 170],
    ["Grand Total", 440],
  ]);
  assert.deepEqual(importedPivot.sourceCapabilities, { sourceBound: true, refreshOnLoadHardenable: true });
  importedPivot.disableRefreshOnLoad();
  assert.equal(importedPivot.refreshPolicy.refreshOnLoad, false);
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const finalZip = await JSZip.loadAsync(new Uint8Array(await final.arrayBuffer()));
  const cacheDefinitionPath = Object.keys(firstZip.files).find((name) => /pivotCache\/pivotCacheDefinition.*\.xml$/i.test(name));
  assert.ok(cacheDefinitionPath, "Native PivotTable must own one cache definition part.");
  const firstCacheXml = await firstZip.file(cacheDefinitionPath).async("text");
  const finalCacheXml = await finalZip.file(cacheDefinitionPath).async("text");
  assert.match(firstCacheXml, /refreshOnLoad="1"/);
  assert.match(finalCacheXml, /refreshOnLoad="0"/);
  assert.equal(
    finalCacheXml.replace(/\srefreshOnLoad="(?:1|true|TRUE|0|false|FALSE)"/, ""),
    firstCacheXml.replace(/\srefreshOnLoad="(?:1|true|TRUE|0|false|FALSE)"/, ""),
  );
  for (const name of Object.keys(firstZip.files).filter((name) => !firstZip.files[name].dir && name !== cacheDefinitionPath)) {
    assert.deepEqual(
      await finalZip.file(name).async("uint8array"),
      await firstZip.file(name).async("uint8array"),
      `Only ${cacheDefinitionPath} may change during PivotTable refresh-on-load hardening (${name}).`,
    );
  }
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  const roundTripPivot = roundTrip.worksheets.getItem(summary.name).pivotTables.items[0];
  const roundTripDatePivot = roundTrip.worksheets.getItem("Date Summary").pivotTables.items[0];
  assert.deepEqual(roundTripPivot.computedValues(), pivot.computedValues());
  assert.equal(roundTripPivot.refreshPolicy.refreshOnLoad, false);
  assert.deepEqual(roundTripPivot.sourceCapabilities, { sourceBound: true, refreshOnLoadHardenable: false });
  assert.deepEqual(roundTripDatePivot.filters, importedDatePivot.filters);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, preview };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "officekit-pivot-table-workflow.xlsx");
  const result = await createPivotTableWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
