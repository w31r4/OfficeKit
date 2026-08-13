import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "office-kit";

const NUMBER_FORMAT = "#,##0.000;[Red](#,##0.000);-";
const HEADER_FILL = "#17324D";
const SECTION_FILL = "#2F5D7C";
const INPUT_FILL = "#FFF2CC";

function assertClose(actual, expected, tolerance = 1e-9) {
  assert.equal(typeof actual, "number");
  assert.ok(Math.abs(actual - expected) <= tolerance, `${actual} should be within ${tolerance} of ${expected}`);
}

function styleTitle(sheet, range, title) {
  sheet.getRange(range).values = [[title, null, null]];
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

export function buildStatisticalAnalysisWorkbook() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const data = workbook.worksheets.add("Data");
  const analysis = workbook.worksheets.add("Analysis");
  const checks = workbook.worksheets.add("Checks");
  for (const sheet of [data, analysis, checks]) sheet.showGridLines = false;

  styleTitle(data, "A1:C1", "Campaign Performance Source Data");
  data.getRange("A3:C9").values = [
    ["Month", "Spend", "Revenue"],
    ["Jan", 10, 22],
    ["Feb", 20, 38],
    ["Mar", 30, 63],
    ["Apr", 40, 77],
    ["May", 50, 102],
    ["Jun", 60, 118],
  ];
  styleHeader(data, "A3:C3");
  data.getRange("B4:C9").format = { fill: INPUT_FILL, font: { color: "#0000FF" }, numberFormat: "$#,##0" };
  data.getRange("A1:A9").format.columnWidthPx = 116;
  data.getRange("B1:C9").format.columnWidthPx = 128;
  data.freezePanes.freezeRows(3);

  styleTitle(analysis, "A1:C1", "Statistical Relationship Analysis");
  analysis.getRange("A3:C3").values = [["Metric", "Spend", "Revenue"]];
  analysis.getRange("A4:C8").values = [
    ["Average", null, null],
    ["Sample variance", null, null],
    ["Sample standard deviation", null, null],
    ["Population variance", null, null],
    ["Population standard deviation", null, null],
  ];
  styleHeader(analysis, "A3:C3");
  analysis.getRange("B4:C8").formulas = [
    ["=AVERAGE('Data'!$B$4:$B$9)", "=AVERAGE('Data'!$C$4:$C$9)"],
    ["=VAR.S('Data'!$B$4:$B$9)", "=VAR.S('Data'!$C$4:$C$9)"],
    ["=STDEV.S('Data'!$B$4:$B$9)", "=STDEV.S('Data'!$C$4:$C$9)"],
    ["=VAR.P('Data'!$B$4:$B$9)", "=VAR.P('Data'!$C$4:$C$9)"],
    ["=STDEV.P('Data'!$B$4:$B$9)", "=STDEV.P('Data'!$C$4:$C$9)"],
  ];
  analysis.getRange("A10:C10").values = [["Relationship metric", "Value", "Interpretation"]];
  analysis.getRange("A11:C13").values = [
    ["Pearson correlation", null, "Direction and strength"],
    ["Sample covariance", null, "Sample co-movement"],
    ["Population covariance", null, "Population co-movement"],
  ];
  analysis.getRange("B11:B13").formulas = [
    ["=CORREL('Data'!$B$4:$B$9,'Data'!$C$4:$C$9)"],
    ["=COVARIANCE.S('Data'!$B$4:$B$9,'Data'!$C$4:$C$9)"],
    ["=COVARIANCE.P('Data'!$B$4:$B$9,'Data'!$C$4:$C$9)"],
  ];
  styleHeader(analysis, "A10:C10");
  analysis.getRange("B4:C8").format = { fill: "#F0FDF4", font: { bold: true, color: "#008000" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("B11:B13").format = { fill: "#F0FDF4", font: { bold: true, color: "#008000" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("A1:A13").format.columnWidthPx = 230;
  analysis.getRange("B1:B13").format.columnWidthPx = 130;
  analysis.getRange("C1:C13").format.columnWidthPx = 210;
  analysis.freezePanes.freezeRows(3);

  styleTitle(checks, "A1:E1", "Independent Model Checks");
  checks.getRange("A3:E3").values = [["Check", "Actual", "Expected / minimum", "Difference", "Status"]];
  checks.getRange("A4:E7").values = [
    ["Correlation is strongly positive", null, 0.99, null, null],
    ["Sample covariance", null, 686, null, null],
    ["Spend sample variance", null, 350, null, null],
    ["Numeric source rows", null, 6, null, null],
  ];
  styleHeader(checks, "A3:E3");
  checks.getRange("B4:B7").formulas = [
    ["='Analysis'!$B$11"],
    ["='Analysis'!$B$12"],
    ["='Analysis'!$B$5"],
    ["=COUNT('Data'!$B$4:$B$9)"],
  ];
  checks.getRange("D4:D7").formulas = [["=B4-C4"], ["=B5-C5"], ["=B6-C6"], ["=B7-C7"]];
  checks.getRange("E4:E7").formulas = [
    ["=IF(B4>=C4,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D5)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D6)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(D7=0,\"OK\",\"CHECK\")"],
  ];
  checks.getRange("B4:D7").format.numberFormat = NUMBER_FORMAT;
  checks.getRange("B4:B7").format.font = { color: "#008000" };
  checks.getRange("E4:E7").format = { fill: "#DCFCE7", font: { bold: true, color: "#166534" }, alignment: { horizontal: "center" } };
  checks.getRange("A1:A7").format.columnWidthPx = 230;
  checks.getRange("B1:D7").format.columnWidthPx = 130;
  checks.getRange("E1:E7").format.columnWidthPx = 82;
  checks.freezePanes.freezeRows(3);

  workbook.worksheets.setActiveWorksheet("Analysis");
  workbook.recalculate();
  return workbook;
}

export async function createStatisticalAnalysisWorkbook(outputPath) {
  const workbook = buildStatisticalAnalysisWorkbook();
  const analysis = workbook.worksheets.getItem("Analysis");
  assertClose(analysis.getRange("B5").values[0][0], 350);
  assertClose(analysis.getRange("B6").values[0][0], Math.sqrt(350));
  assertClose(analysis.getRange("B12").values[0][0], 686);
  assertClose(analysis.getRange("B13").values[0][0], 3430 / 6);
  assertClose(analysis.getRange("B11").values[0][0], 3430 / Math.sqrt(1750 * 6754));
  assert.deepEqual(workbook.worksheets.getItem("Checks").getRange("E4:E7").values, [["OK"], ["OK"], ["OK"], ["OK"]]);

  const inspection = workbook.inspect({ kind: "workbook,sheet,formula", sheetName: "Analysis", range: "A1:C13", maxChars: 12_000 });
  assert.match(inspection.ndjson, /CORREL/);
  assert.match(inspection.ndjson, /COVARIANCE\.S/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = await workbook.render({ sheetName: "Analysis", range: "A1:C13", autoCrop: "all", format: "svg" });
  assert.match(await previewSvg.text(), /Statistical Relationship Analysis/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  assert.equal(imported.worksheets.getItem("Analysis").getRange("B11").formulas[0][0], "=CORREL('Data'!$B$4:$B$9,'Data'!$C$4:$C$9)");
  assert.deepEqual(imported.worksheets.getItem("Checks").getRange("E4:E7").values, [["OK"], ["OK"], ["OK"], ["OK"]]);
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  roundTrip.recalculate();
  assertClose(roundTrip.worksheets.getItem("Analysis").getRange("B12").values[0][0], 686);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, previewSvg };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "officekit-statistical-analysis-workflow.xlsx");
  const result = await createStatisticalAnalysisWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
