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

export function buildExponentialGrowthWorkbook() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const data = workbook.worksheets.add("Data");
  const analysis = workbook.worksheets.add("Analysis");
  const checks = workbook.worksheets.add("Checks");
  for (const sheet of [data, analysis, checks]) sheet.showGridLines = false;

  styleTitle(data, "A1:C1", "Observed Growth Source Data");
  data.getRange("A3:C9").values = [
    ["Period", "Observed units", "Natural log"],
    [2, 6, null],
    [3, 11, null],
    [4, 18, null],
    [5, 33, null],
    [6, 54, null],
    [7, 91, null],
  ];
  data.getRange("C4:C9").formulas = Array.from({ length: 6 }, (_, index) => [`=LN(B${index + 4})`]);
  styleHeader(data, "A3:C3");
  data.getRange("A4:B9").format = { fill: INPUT_FILL, font: { color: "#0000FF" }, numberFormat: NUMBER_FORMAT };
  data.getRange("C4:C9").format = { fill: "#F0FDF4", font: { color: "#008000" }, numberFormat: NUMBER_FORMAT };
  data.getRange("E3:F6").values = [
    ["Forecast horizon", "Period"],
    ["Point 1", 8],
    ["Point 2", 9],
    ["Point 3", 10],
  ];
  styleHeader(data, "E3:F3");
  data.getRange("F4:F6").format = { fill: INPUT_FILL, font: { color: "#0000FF" }, numberFormat: "0" };
  data.getRange("A1:A9").format.columnWidthPx = 116;
  data.getRange("B1:C9").format.columnWidthPx = 132;
  data.getRange("E1:E6").format.columnWidthPx = 150;
  data.getRange("F1:F6").format.columnWidthPx = 100;
  data.freezePanes.freezeRows(3);

  styleTitle(analysis, "A1:I1", "Exponential Growth Analysis");
  analysis.getRange("A3:B3").values = [["Model metric", "Value"]];
  analysis.getRange("A4:B9").values = [
    ["Multiplier m", null],
    ["Base b", null],
    ["Log-space slope", null],
    ["Log-space intercept", null],
    ["Log-space R-squared", null],
    ["Point 1 forecast", null],
  ];
  analysis.getRange("B4:B9").formulas = [
    ["=EXP(B6)"],
    ["=EXP(B7)"],
    ["=SLOPE('Data'!$C$4:$C$9,'Data'!$A$4:$A$9)"],
    ["=INTERCEPT('Data'!$C$4:$C$9,'Data'!$A$4:$A$9)"],
    ["=RSQ('Data'!$C$4:$C$9,'Data'!$A$4:$A$9)"],
    ["=EXP(B7+B6*'Data'!$F$4)"],
  ];
  styleHeader(analysis, "A3:B3");
  analysis.getRange("D3:F3").values = [["LOGEST statistic", "Multiplier / model", "Base / residual"]];
  analysis.getRange("D4:D8").values = [
    ["Coefficient"],
    ["Log-coefficient standard error"],
    ["R-squared / log standard error"],
    ["F-statistic / degrees freedom"],
    ["Log regression / residual SS"],
  ];
  analysis.getRange("E4").formulas = [["=LOGEST('Data'!$B$4:$B$9,'Data'!$A$4:$A$9,TRUE,TRUE)"]];
  styleHeader(analysis, "D3:F3");
  analysis.getRange("H3:I3").values = [["Forecast period", "GROWTH units"]];
  analysis.getRange("H4:H6").formulas = [["='Data'!$F$4"], ["='Data'!$F$5"], ["='Data'!$F$6"]];
  analysis.getRange("I4").formulas = [["=GROWTH('Data'!$B$4:$B$9,'Data'!$A$4:$A$9,H4:H6)"]];
  styleHeader(analysis, "H3:I3");
  analysis.getRange("B4:B9").format = { fill: "#EFF6FF", font: { bold: true, color: "#1D4ED8" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("E4:F8").format = { fill: "#F5F3FF", font: { bold: true, color: "#6D28D9" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("H4:I6").format = { fill: "#FFF7ED", font: { bold: true, color: "#C2410C" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("A1:A9").format.columnWidthPx = 220;
  analysis.getRange("B1:B9").format.columnWidthPx = 135;
  analysis.getRange("C1:C9").format.columnWidthPx = 18;
  analysis.getRange("D1:D9").format.columnWidthPx = 235;
  analysis.getRange("E1:F9").format.columnWidthPx = 145;
  analysis.getRange("G1:G9").format.columnWidthPx = 18;
  analysis.getRange("H1:I9").format.columnWidthPx = 145;
  analysis.freezePanes.freezeRows(3);

  styleTitle(checks, "A1:E1", "Independent Exponential Model Checks");
  checks.getRange("A3:E3").values = [["Check", "Actual", "Expected", "Difference", "Status"]];
  checks.getRange("A4:E11").values = [
    ["LOGEST multiplier reconciles", null, null, null, null],
    ["LOGEST base reconciles", null, null, null, null],
    ["LOGEST R-squared reconciles", null, null, null, null],
    ["GROWTH point 1 reconciles", null, null, null, null],
    ["GROWTH point 2 reconciles", null, null, null, null],
    ["GROWTH point 3 reconciles", null, null, null, null],
    ["Forecasts are positive", null, 3, null, null],
    ["Numeric source pairs", null, 6, null, null],
  ];
  checks.getRange("B4:B11").formulas = [
    ["='Analysis'!$E$4"],
    ["='Analysis'!$F$4"],
    ["='Analysis'!$E$6"],
    ["='Analysis'!$I$4"],
    ["='Analysis'!$I$5"],
    ["='Analysis'!$I$6"],
    ["=COUNTIF('Analysis'!$I$4:$I$6,\">0\")"],
    ["=COUNT('Data'!$A$4:$A$9)"],
  ];
  checks.getRange("C4:C9").formulas = [
    ["=EXP('Analysis'!$B$6)"],
    ["=EXP('Analysis'!$B$7)"],
    ["='Analysis'!$B$8"],
    ["=EXP('Analysis'!$B$7+'Analysis'!$B$6*'Analysis'!$H$4)"],
    ["=EXP('Analysis'!$B$7+'Analysis'!$B$6*'Analysis'!$H$5)"],
    ["=EXP('Analysis'!$B$7+'Analysis'!$B$6*'Analysis'!$H$6)"],
  ];
  checks.getRange("D4:D11").formulas = Array.from({ length: 8 }, (_, index) => [`=B${index + 4}-C${index + 4}`]);
  checks.getRange("E4:E11").formulas = Array.from({ length: 8 }, (_, index) => [`=IF(ABS(D${index + 4})<0.000001,"OK","CHECK")`]);
  styleHeader(checks, "A3:E3");
  checks.getRange("B4:D11").format.numberFormat = NUMBER_FORMAT;
  checks.getRange("B4:B11").format.font = { color: "#008000" };
  checks.getRange("E4:E11").format = { fill: "#DCFCE7", font: { bold: true, color: "#166534" }, alignment: { horizontal: "center" } };
  checks.getRange("A1:A11").format.columnWidthPx = 260;
  checks.getRange("B1:D11").format.columnWidthPx = 130;
  checks.getRange("E1:E11").format.columnWidthPx = 82;
  checks.freezePanes.freezeRows(3);

  workbook.worksheets.setActiveWorksheet("Analysis");
  workbook.recalculate();
  return workbook;
}

export async function createExponentialGrowthWorkbook(outputPath) {
  const workbook = buildExponentialGrowthWorkbook();
  const analysis = workbook.worksheets.getItem("Analysis");
  assertClose(analysis.getRange("E4").values[0][0], 1.71968316255041);
  assertClose(analysis.getRange("F4").values[0][0], 2.0954450641003);
  assertClose(analysis.getRange("E6").values[0][0], 0.999014535947832);
  assertClose(analysis.getRange("I4").values[0][0], 160.27426439333);
  assertClose(analysis.getRange("I5").values[0][0], 275.620953867361);
  assertClose(analysis.getRange("I6").values[0][0], 473.980713611783);
  assert.equal(analysis.store.get("E4").spillRange, "E4:F8");
  assert.equal(analysis.store.get("I4").spillRange, "I4:I6");
  assert.deepEqual(workbook.worksheets.getItem("Checks").getRange("E4:E11").values, Array.from({ length: 8 }, () => ["OK"]));

  const inspection = workbook.inspect({ kind: "workbook,sheet,formula", sheetName: "Analysis", range: "A1:I9", maxChars: 18_000 });
  assert.match(inspection.ndjson, /LOGEST/);
  assert.match(inspection.ndjson, /GROWTH/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = await workbook.render({ sheetName: "Analysis", range: "A1:I9", autoCrop: "all", format: "svg" });
  assert.match(await previewSvg.text(), /Exponential Growth Analysis/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  const importedAnalysis = imported.worksheets.getItem("Analysis");
  assert.equal(importedAnalysis.getRange("E4").formulas[0][0], "=LOGEST('Data'!$B$4:$B$9,'Data'!$A$4:$A$9,TRUE,TRUE)");
  assert.equal(importedAnalysis.getRange("I4").formulas[0][0], "=GROWTH('Data'!$B$4:$B$9,'Data'!$A$4:$A$9,H4:H6)");
  assert.equal(importedAnalysis.store.get("E4").dynamicArrayRef, "E4:F8");
  assert.equal(importedAnalysis.store.get("I4").dynamicArrayRef, "I4:I6");
  assert.deepEqual(imported.worksheets.getItem("Checks").getRange("E4:E11").values, Array.from({ length: 8 }, () => ["OK"]));
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  roundTrip.recalculate();
  assertClose(roundTrip.worksheets.getItem("Analysis").getRange("E4").values[0][0], 1.71968316255041);
  assertClose(roundTrip.worksheets.getItem("Analysis").getRange("I6").values[0][0], 473.980713611783);
  assert.deepEqual(roundTrip.worksheets.getItem("Checks").getRange("E4:E11").values, Array.from({ length: 8 }, () => ["OK"]));

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, previewSvg };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "officekit-exponential-growth-workflow.xlsx");
  const result = await createExponentialGrowthWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
