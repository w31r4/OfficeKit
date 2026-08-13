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
  data.getRange("E3:F6").values = [
    ["Forecast horizon", "Spend"],
    ["Point 1", 70],
    ["Point 2", 80],
    ["Point 3", 90],
  ];
  styleHeader(data, "E3:F3");
  data.getRange("F4:F6").format = { fill: INPUT_FILL, font: { color: "#0000FF" }, numberFormat: "$#,##0" };
  data.getRange("E1:E6").format.columnWidthPx = 150;
  data.getRange("F1:F6").format.columnWidthPx = 110;
  data.freezePanes.freezeRows(3);

  styleTitle(analysis, "A1:I1", "Statistical Relationship Analysis");
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
  analysis.getRange("A15:C15").values = [["Linear regression metric", "Value", "Interpretation"]];
  analysis.getRange("A16:C21").values = [
    ["Slope", null, "Revenue change per spend unit"],
    ["Intercept", null, "Estimated revenue at zero spend"],
    ["R-squared", null, "Explained variance proportion"],
    ["Standard error", null, "Prediction residual scale"],
    ["Forecast spend", null, "Visible source assumption"],
    ["Forecast revenue", null, "Linear point estimate"],
  ];
  analysis.getRange("B16:B21").formulas = [
    ["=SLOPE('Data'!$C$4:$C$9,'Data'!$B$4:$B$9)"],
    ["=INTERCEPT('Data'!$C$4:$C$9,'Data'!$B$4:$B$9)"],
    ["=RSQ('Data'!$C$4:$C$9,'Data'!$B$4:$B$9)"],
    ["=STEYX('Data'!$C$4:$C$9,'Data'!$B$4:$B$9)"],
    ["='Data'!$F$4"],
    ["=FORECAST.LINEAR(B20,'Data'!$C$4:$C$9,'Data'!$B$4:$B$9)"],
  ];
  styleHeader(analysis, "A15:C15");
  analysis.getRange("D15:F15").values = [["LINEST statistic", "Slope / model", "Intercept / residual"]];
  analysis.getRange("D16:D20").values = [
    ["Coefficient"],
    ["Coefficient standard error"],
    ["R-squared / standard error"],
    ["F-statistic / degrees freedom"],
    ["Regression / residual SS"],
  ];
  analysis.getRange("E16").formulas = [["=LINEST('Data'!$C$4:$C$9,'Data'!$B$4:$B$9,TRUE,TRUE)"]];
  styleHeader(analysis, "D15:F15");
  analysis.getRange("H15:I15").values = [["Forecast spend", "TREND revenue"]];
  analysis.getRange("H16:H18").formulas = [["='Data'!$F$4"], ["='Data'!$F$5"], ["='Data'!$F$6"]];
  analysis.getRange("I16").formulas = [["=TREND('Data'!$C$4:$C$9,'Data'!$B$4:$B$9,H16:H18)"]];
  styleHeader(analysis, "H15:I15");
  analysis.getRange("B4:C8").format = { fill: "#F0FDF4", font: { bold: true, color: "#008000" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("B11:B13").format = { fill: "#F0FDF4", font: { bold: true, color: "#008000" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("B16:B21").format = { fill: "#EFF6FF", font: { bold: true, color: "#1D4ED8" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("E16:F20").format = { fill: "#F5F3FF", font: { bold: true, color: "#6D28D9" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("H16:I18").format = { fill: "#FFF7ED", font: { bold: true, color: "#C2410C" }, numberFormat: NUMBER_FORMAT };
  analysis.getRange("A1:A21").format.columnWidthPx = 230;
  analysis.getRange("B1:B21").format.columnWidthPx = 130;
  analysis.getRange("C1:C21").format.columnWidthPx = 210;
  analysis.getRange("D1:D21").format.columnWidthPx = 210;
  analysis.getRange("E1:F21").format.columnWidthPx = 140;
  analysis.getRange("G1:G21").format.columnWidthPx = 18;
  analysis.getRange("H1:I21").format.columnWidthPx = 140;
  analysis.freezePanes.freezeRows(3);

  styleTitle(checks, "A1:E1", "Independent Model Checks");
  checks.getRange("A3:E3").values = [["Check", "Actual", "Expected / minimum", "Difference", "Status"]];
  checks.getRange("A4:E18").values = [
    ["Correlation is strongly positive", null, 0.99, null, null],
    ["Sample covariance", null, 686, null, null],
    ["Spend sample variance", null, 350, null, null],
    ["Numeric source rows", null, 6, null, null],
    ["Regression slope", null, 1.96, null, null],
    ["Regression intercept", null, 1.4, null, null],
    ["R-squared equals correlation squared", null, null, null, null],
    ["Forecast reconciles to line equation", null, null, null, null],
    ["Standard error is nonnegative", null, 0, null, null],
    ["LINEST slope reconciles", null, null, null, null],
    ["LINEST intercept reconciles", null, null, null, null],
    ["LINEST R-squared reconciles", null, null, null, null],
    ["TREND point 1 reconciles", null, null, null, null],
    ["TREND point 2 reconciles", null, null, null, null],
    ["TREND point 3 reconciles", null, null, null, null],
  ];
  styleHeader(checks, "A3:E3");
  checks.getRange("B4:B18").formulas = [
    ["='Analysis'!$B$11"],
    ["='Analysis'!$B$12"],
    ["='Analysis'!$B$5"],
    ["=COUNT('Data'!$B$4:$B$9)"],
    ["='Analysis'!$B$16"],
    ["='Analysis'!$B$17"],
    ["='Analysis'!$B$18"],
    ["='Analysis'!$B$21"],
    ["='Analysis'!$B$19"],
    ["='Analysis'!$E$16"],
    ["='Analysis'!$F$16"],
    ["='Analysis'!$E$18"],
    ["='Analysis'!$I$16"],
    ["='Analysis'!$I$17"],
    ["='Analysis'!$I$18"],
  ];
  checks.getRange("C10:C11").formulas = [["='Analysis'!$B$11^2"], ["='Analysis'!$B$17+'Analysis'!$B$16*'Analysis'!$B$20"]];
  checks.getRange("C13:C15").formulas = [["='Analysis'!$B$16"], ["='Analysis'!$B$17"], ["='Analysis'!$B$18"]];
  checks.getRange("C16:C18").formulas = [["='Analysis'!$B$17+'Analysis'!$B$16*'Analysis'!$H$16"], ["='Analysis'!$B$17+'Analysis'!$B$16*'Analysis'!$H$17"], ["='Analysis'!$B$17+'Analysis'!$B$16*'Analysis'!$H$18"]];
  checks.getRange("D4:D18").formulas = Array.from({ length: 15 }, (_, index) => [`=B${index + 4}-C${index + 4}`]);
  checks.getRange("E4:E18").formulas = [
    ["=IF(B4>=C4,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D5)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D6)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(D7=0,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D8)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D9)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D10)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D11)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(B12>=C12,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D13)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D14)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D15)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D16)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D17)<0.000001,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D18)<0.000001,\"OK\",\"CHECK\")"],
  ];
  checks.getRange("B4:D18").format.numberFormat = NUMBER_FORMAT;
  checks.getRange("B4:B18").format.font = { color: "#008000" };
  checks.getRange("E4:E18").format = { fill: "#DCFCE7", font: { bold: true, color: "#166534" }, alignment: { horizontal: "center" } };
  checks.getRange("A1:A18").format.columnWidthPx = 260;
  checks.getRange("B1:D18").format.columnWidthPx = 130;
  checks.getRange("E1:E18").format.columnWidthPx = 82;
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
  assertClose(analysis.getRange("B16").values[0][0], 1.96);
  assertClose(analysis.getRange("B17").values[0][0], 1.4);
  assertClose(analysis.getRange("B18").values[0][0], (3430 * 3430) / (1750 * 6754));
  assertClose(analysis.getRange("B19").values[0][0], 2.79284800875378);
  assertClose(analysis.getRange("B21").values[0][0], 138.6);
  assertClose(analysis.getRange("E16").values[0][0], 1.96);
  assertClose(analysis.getRange("F16").values[0][0], 1.4);
  assertClose(analysis.getRange("E18").values[0][0], (3430 * 3430) / (1750 * 6754));
  assertClose(analysis.getRange("F18").values[0][0], 2.79284800875378);
  assert.deepEqual(analysis.getRange("I16:I18").values, [[138.6], [158.2], [177.8]]);
  assert.equal(analysis.store.get("I16").spillRange, "I16:I18");
  assert.equal(analysis.store.get("E16").spillRange, "E16:F20");
  assert.deepEqual(workbook.worksheets.getItem("Checks").getRange("E4:E18").values, Array.from({ length: 15 }, () => ["OK"]));

  const inspection = workbook.inspect({ kind: "workbook,sheet,formula", sheetName: "Analysis", range: "A1:I21", maxChars: 24_000 });
  assert.match(inspection.ndjson, /CORREL/);
  assert.match(inspection.ndjson, /COVARIANCE\.S/);
  assert.match(inspection.ndjson, /FORECAST\.LINEAR/);
  assert.match(inspection.ndjson, /LINEST/);
  assert.match(inspection.ndjson, /TREND/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = await workbook.render({ sheetName: "Analysis", range: "A1:I21", autoCrop: "all", format: "svg" });
  assert.match(await previewSvg.text(), /Statistical Relationship Analysis/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  assert.equal(imported.worksheets.getItem("Analysis").getRange("B11").formulas[0][0], "=CORREL('Data'!$B$4:$B$9,'Data'!$C$4:$C$9)");
  assert.equal(imported.worksheets.getItem("Analysis").getRange("B21").formulas[0][0], "=FORECAST.LINEAR(B20,'Data'!$C$4:$C$9,'Data'!$B$4:$B$9)");
  assert.equal(imported.worksheets.getItem("Analysis").getRange("E16").formulas[0][0], "=LINEST('Data'!$C$4:$C$9,'Data'!$B$4:$B$9,TRUE,TRUE)");
  assert.equal(imported.worksheets.getItem("Analysis").getRange("I16").formulas[0][0], "=TREND('Data'!$C$4:$C$9,'Data'!$B$4:$B$9,H16:H18)");
  assert.equal(imported.worksheets.getItem("Analysis").store.get("E16").formulaType, "dynamicArray");
  assert.equal(imported.worksheets.getItem("Analysis").store.get("E16").dynamicArrayRef, "E16:F20");
  assert.equal(imported.worksheets.getItem("Analysis").store.get("I16").dynamicArrayRef, "I16:I18");
  assert.deepEqual(imported.worksheets.getItem("Checks").getRange("E4:E18").values, Array.from({ length: 15 }, () => ["OK"]));
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  roundTrip.recalculate();
  assertClose(roundTrip.worksheets.getItem("Analysis").getRange("B12").values[0][0], 686);
  assertClose(roundTrip.worksheets.getItem("Analysis").getRange("B21").values[0][0], 138.6);
  assertClose(roundTrip.worksheets.getItem("Analysis").getRange("E16").values[0][0], 1.96);
  assertClose(roundTrip.worksheets.getItem("Analysis").getRange("F18").values[0][0], 2.79284800875378);
  assert.deepEqual(roundTrip.worksheets.getItem("Analysis").getRange("I16:I18").values, [[138.6], [158.2], [177.8]]);

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
