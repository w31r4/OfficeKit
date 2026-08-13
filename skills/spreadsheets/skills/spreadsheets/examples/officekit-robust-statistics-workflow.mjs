import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "office-kit";

const HEADER_FILL = "#17324D";
const SECTION_FILL = "#2F5D7C";

function styleTitle(sheet, range, title) {
  sheet.getRange(range).values = [[title, null, null, null]];
  sheet.getRange(range).merge();
  sheet.getRange(range).format = {
    fill: HEADER_FILL,
    font: { bold: true, color: "#FFFFFF" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 30,
  };
}

function styleHeader(sheet, range) {
  sheet.getRange(range).format = {
    fill: SECTION_FILL,
    font: { bold: true, color: "#FFFFFF" },
    alignment: { horizontal: "center", vertical: "center" },
  };
}

export function buildRobustStatisticsWorkbook() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const data = workbook.worksheets.add("Data");
  const analysis = workbook.worksheets.add("Analysis");
  const checks = workbook.worksheets.add("Checks");
  for (const sheet of [data, analysis, checks]) sheet.showGridLines = false;

  styleTitle(data, "A1:D1", "Robust Statistics Source Data");
  data.getRange("A3:B10").values = [
    ["Observation", "Value"],
    ["Sample 1", 1],
    ["Sample 2", 2],
    ["Sample 3", 2],
    ["Sample 4", 3],
    ["Sample 5", 3],
    ["Sample 6", 4],
    ["Outlier", 100],
  ];
  styleHeader(data, "A3:B3");
  data.getRange("B4:B10").format = { fill: "#FFF2CC", font: { color: "#0000FF" }, numberFormat: "#,##0.000" };
  data.getRange("A1:A10").format.columnWidthPx = 145;
  data.getRange("B1:B10").format.columnWidthPx = 120;
  data.freezePanes.freezeRows(3);

  styleTitle(analysis, "A1:F1", "Robust and Order Statistics");
  analysis.getRange("A3:B3").values = [["Metric", "Result"]];
  analysis.getRange("A4:B8").values = [
    ["Average descending rank for tied value 2", null],
    ["Exclusive median percentile", null],
    ["Exclusive first quartile", null],
    ["30% trimmed mean", null],
    ["First mode", null],
  ];
  analysis.getRange("B4:B8").formulas = [
    ["=_xlfn.RANK.AVG(2,'Data'!$B$4:$B$10,0)"],
    ["=_xlfn.PERCENTILE.EXC('Data'!$B$4:$B$10,0.5)"],
    ["=_xlfn.QUARTILE.EXC('Data'!$B$4:$B$10,1)"],
    ["=TRIMMEAN('Data'!$B$4:$B$10,0.3)"],
    ["=MODE.SNGL('Data'!$B$4:$B$10)"],
  ];
  styleHeader(analysis, "A3:B3");
  analysis.getRange("D3:E3").values = [["All tied modes", "Frequency"]];
  analysis.getRange("D4").formulas = [["=_xlfn.MODE.MULT('Data'!$B$4:$B$10)"]];
  analysis.getRange("E4:E5").formulas = [
    ["=COUNTIF('Data'!$B$4:$B$10,D4)"],
    ["=COUNTIF('Data'!$B$4:$B$10,D5)"],
  ];
  styleHeader(analysis, "D3:E3");
  analysis.getRange("B4:B8").format = { fill: "#EFF6FF", font: { bold: true, color: "#1D4ED8" }, numberFormat: "#,##0.000" };
  analysis.getRange("D4:E5").format = { fill: "#F5F3FF", font: { bold: true, color: "#6D28D9" }, numberFormat: "#,##0.000" };
  analysis.getRange("A1:A8").format.columnWidthPx = 270;
  analysis.getRange("B1:B8").format.columnWidthPx = 125;
  analysis.getRange("C1:C8").format.columnWidthPx = 22;
  analysis.getRange("D1:D8").format.columnWidthPx = 145;
  analysis.getRange("E1:E8").format.columnWidthPx = 105;
  analysis.freezePanes.freezeRows(3);

  styleTitle(checks, "A1:E1", "Independent Robust Statistics Checks");
  checks.getRange("A3:E3").values = [["Check", "Actual", "Expected", "Difference", "Status"]];
  checks.getRange("A4:E10").values = [
    ["Average rank reconciles", null, null, null, null],
    ["Exclusive median reconciles", null, null, null, null],
    ["Exclusive first quartile reconciles", null, null, null, null],
    ["Trimmed mean removes both tails", null, null, null, null],
    ["First mode reconciles", null, null, null, null],
    ["Both modes have equal frequency", null, null, null, null],
    ["MODE.MULT spills two rows", null, 2, null, null],
  ];
  checks.getRange("B4:B10").formulas = [
    ["='Analysis'!$B$4"], ["='Analysis'!$B$5"], ["='Analysis'!$B$6"], ["='Analysis'!$B$7"],
    ["='Analysis'!$D$4"], ["='Analysis'!$E$4-'Analysis'!$E$5"], ["=COUNT('Analysis'!$D$4:$D$5)"],
  ];
  checks.getRange("C4:C9").formulas = [
    ["=RANK.EQ(2,'Data'!$B$4:$B$10,0)+(COUNTIF('Data'!$B$4:$B$10,2)-1)/2"],
    ["=MEDIAN('Data'!$B$4:$B$10)"],
    ["=SMALL('Data'!$B$4:$B$10,2)"],
    ["=AVERAGE(2,2,3,3,4)"],
    ["=MODE.SNGL('Data'!$B$4:$B$10)"],
    ["=0"],
  ];
  checks.getRange("D4:D10").formulas = Array.from({ length: 7 }, (_, index) => [`=B${index + 4}-C${index + 4}`]);
  checks.getRange("E4:E10").formulas = Array.from({ length: 7 }, (_, index) => [`=IF(ABS(D${index + 4})<0.000001,"OK","CHECK")`]);
  styleHeader(checks, "A3:E3");
  checks.getRange("B4:D10").format.numberFormat = "#,##0.000";
  checks.getRange("E4:E10").format = { fill: "#DCFCE7", font: { bold: true, color: "#166534" }, alignment: { horizontal: "center" } };
  checks.getRange("A1:A10").format.columnWidthPx = 265;
  checks.getRange("B1:D10").format.columnWidthPx = 120;
  checks.getRange("E1:E10").format.columnWidthPx = 82;
  checks.freezePanes.freezeRows(3);

  workbook.worksheets.setActiveWorksheet("Analysis");
  workbook.recalculate();
  return workbook;
}

export async function createRobustStatisticsWorkbook(outputPath) {
  const workbook = buildRobustStatisticsWorkbook();
  const analysis = workbook.worksheets.getItem("Analysis");
  assert.deepEqual(analysis.getRange("B4:B8").values.flat(), [5.5, 3, 2, 2.8, 2]);
  assert.deepEqual(analysis.getRange("D4:E5").values, [[2, 2], [3, 2]]);
  assert.equal(analysis.store.get("D4").spillRange, "D4:D5");
  assert.deepEqual(workbook.worksheets.getItem("Checks").getRange("E4:E10").values, Array.from({ length: 7 }, () => ["OK"]));

  const inspection = workbook.inspect({ kind: "workbook,sheet,formula", sheetName: "Analysis", range: "A1:E8", maxChars: 16_000 });
  assert.match(inspection.ndjson, /RANK\.AVG/);
  assert.match(inspection.ndjson, /MODE\.MULT/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = await workbook.render({ sheetName: "Analysis", range: "A1:E8", autoCrop: "all", format: "svg" });
  assert.match(await previewSvg.text(), /Robust and Order Statistics/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  assert.deepEqual(imported.worksheets.getItem("Analysis").getRange("D4:D5").values, [[2], [3]]);
  assert.equal(imported.worksheets.getItem("Analysis").store.get("D4").dynamicArrayRef, "D4:D5");
  assert.deepEqual(imported.worksheets.getItem("Checks").getRange("E4:E10").values, Array.from({ length: 7 }, () => ["OK"]));
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  roundTrip.recalculate();
  assert.deepEqual(roundTrip.worksheets.getItem("Analysis").getRange("B4:B8").values.flat(), [5.5, 3, 2, 2.8, 2]);
  assert.deepEqual(roundTrip.worksheets.getItem("Analysis").getRange("D4:D5").values, [[2], [3]]);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, previewSvg };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "officekit-robust-statistics-workflow.xlsx");
  const result = await createRobustStatisticsWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
