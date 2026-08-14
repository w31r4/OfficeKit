import assert from "node:assert/strict";
import JSZip from "jszip";

import { FileBlob, SpreadsheetFile, Workbook, verifyArtifact } from "../src/index.mjs";
import { formatSpreadsheetDisplayValue } from "../src/spreadsheet/ooxml-styles.mjs";

const PNG_DATA_URL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAABFaP0WAAAADUlEQVR42mNk+M/wHwAF/gL+3c5GAAAAAElFTkSuQmCC";

assert.equal(formatSpreadsheetDisplayValue(-8884.878867834168, { numberFormat: "$#,##0;[Red]($#,##0);-" }), "($8,885)");

const workbook = Workbook.create({ dateSystem: "1900" });
const sheet = workbook.worksheets.add("Summary");
sheet.showGridLines = false;
sheet.getRange("A1:E4").values = [
  ["Month", "Revenue", "Cost", "Status", "Date"],
  ["Jan", 100, 60, "Done", new Date("2026-01-15T00:00:00.000Z")],
  ["Feb", 120, 70, "Review", new Date("2026-02-15T00:00:00.000Z")],
  ["Mar", 150, 90, "Planned", new Date("2026-03-15T00:00:00.000Z")],
];
sheet.getRange("F1:F4").values = [["Margin"], [null], [null], [null]];
sheet.getRange("F2:F4").formulas = [
  ["=(B2-C2)/B2"],
  ["=(B3-C3)/B3"],
  ["=(B4-C4)/B4"],
];
sheet.getRange("A1:F1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF" },
  border: { bottom: { style: "double", color: "#38BDF8" } },
  alignment: { horizontal: "center", vertical: "center" },
};
sheet.getRange("B2:C4").format = { numberFormat: "$#,##0.00" };
sheet.getRange("E2:E4").format = { numberFormat: "yyyy-mm-dd" };
sheet.getRange("F2:F4").format = {
  numberFormat: "0.0%",
  protection: { locked: false, hidden: true },
};
sheet.getRange("A1:A6").format.columnWidthPx = 96;
sheet.getRange("B1:F6").format.columnWidthPx = 84;
sheet.getRange("A1:F1").format.rowHeightPx = 28;
sheet.getRange("A8:F8").format.rowHidden = true;
sheet.getRange("A6:F6").values = [["Quarter summary", null, null, null, null, null]];
sheet.getRange("A6:F6").merge();
sheet.freezePanes.freezeRows(1);
sheet.freezePanes.freezeColumns(1);
sheet.protection = {
  allow: ["selectLockedCells", "selectUnlockedCells", "sort", "autoFilter"],
};
const copiedProtection = sheet.protection;
copiedProtection.allow.push("formatCells");
assert.deepEqual(sheet.protection, {
  enabled: true,
  allow: ["selectLockedCells", "selectUnlockedCells", "sort", "autoFilter"],
});
assert.throws(() => { sheet.protection = { allow: ["unknownOperation"] }; }, /unsupported worksheet protection operation/i);
assert.throws(() => { sheet.protection = { password: "secret" }; }, /unsupported field.*password.*intentionally not accepted/i);
assert.throws(() => { sheet.protection = { enabled: false, allow: ["sort"] }; }, /disabled worksheet protection cannot declare allowed operations/i);

const table = sheet.tables.add("A1:F4", true, "SummaryTable");
table.style = "TableStyleMedium4";
table.showRowStripes = true;

sheet.getRange("D2:D4").dataValidation = {
  rule: {
    type: "list",
    values: ["Planned", "Review", "Done"],
    allowBlank: false,
    showInputMessage: true,
    promptTitle: "Choose a status",
    prompt: "Use one approved workflow state.",
    showErrorMessage: true,
    errorTitle: "Invalid status",
    error: "Choose a value from the list.",
    errorStyle: "warning",
    showDropdown: true,
  },
};
sheet.dataValidations.add({
  range: "B2:B4",
  rule: { type: "whole", operator: "between", formula1: "0", formula2: "1000" },
});
sheet.dataValidations.add({
  range: "C2:C4",
  rule: {
    type: "custom",
    formula1: "=C2<=B2",
    allowBlank: true,
    showErrorMessage: true,
    errorTitle: "Cost exceeds revenue",
    error: "Enter a cost no greater than revenue.",
    errorStyle: "stop",
  },
});
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "whole", formula1: "0", showDropdown: true } }),
  /showDropdown is valid only for list rules/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: ["North, East"] } }),
  /inline values must be non-empty and cannot contain commas or control characters/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: [""] } }),
  /inline values must be non-empty/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: ["Ready\u007F"] } }),
  /control characters/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "custom", formula1: "=A2<>\"\"", errorStyle: "retry" } }),
  /errorStyle must be one of stop, warning, information/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: ["Ready"], promptTitle: "x".repeat(33) } }),
  /promptTitle must be at most 32 characters/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: ["Ready"], imeMode: "fullAlpha" } }),
  /unsupported field: imeMode/i,
);
sheet.getRange("F2:F4").conditionalFormats.add("cellIs", {
  operator: "greaterThan",
  formula: "0.4",
  format: { fill: "#DCFCE7", font: { bold: true, color: "#166534" } },
});
sheet.getRange("B2:B4").conditionalFormats.add("expression", {
  formula: "B2>C2",
  format: { fill: "#DBEAFE" },
});
sheet.getRange("D2:D4").conditionalFormats.add("containsText", {
  text: "Review",
  formula: "NOT(ISERROR(SEARCH(\"Review\",D2)))",
  format: { fill: "#FEF3C7" },
});
sheet.getRange("C2:C4").conditionalFormats.addColorScale({
  colors: ["#FEE2E2", "#FEF3C7", "#22C55E"],
});
sheet.getRange("B2:B4").conditionalFormats.add("dataBar", {
  color: "#2563EB",
  thresholds: ["min", "max"],
  showValue: false,
});
sheet.getRange("C2:C4").conditionalFormats.add("iconSet", {
  iconSet: "3Arrows",
  thresholds: [0, "50%", { type: "percent", value: 80 }],
  reverse: true,
});
assert.throws(
  () => sheet.getRange("B2:B4").conditionalFormats.add("dataBar", { gradient: false }),
  /gradient=false requires the x14 solid-data-bar extension/i,
);
assert.throws(
  () => sheet.getRange("C2:C4").conditionalFormats.add("iconSet", { iconSet: "3Triangles" }),
  /requires the x14 extension namespace/i,
);
assert.throws(
  () => sheet.getRange("C2:C4").conditionalFormats.add("iconSet", { iconSet: "3Arrows", showValue: "false" }),
  /showValue must be a boolean/i,
);

const bar = sheet.charts.add("bar", sheet.getRange("A1:C4"));
bar.name = "Bar chart";
bar.title = "Revenue and cost";
bar.setAccessibilityMetadata({ title: "Quarterly revenue and cost chart", description: "Revenue and cost both rise from January through March." });
bar.setPosition("H1", "L10");
const line = sheet.charts.add("line", sheet.getRange("A1:C4"));
line.name = "Line chart";
line.title = "Revenue trend";
line.series.items[0].trendlines = [
  {
    type: "linear",
    name: "Revenue projection",
    forward: 0.5,
    backward: 0.5,
    intercept: 0,
    displayEquation: true,
    displayRSquared: true,
    line: { fill: "#7C3AED", style: "dashed", width: 1.5 },
  },
  { type: "movingAverage", name: "Revenue moving average", period: 2 },
  { type: "polynomial", name: "Revenue curve", order: 2 },
];
line.series.items[0].errorBars = {
  type: "percentage",
  value: 10,
  endStyle: "noCap",
  line: { fill: "#DC2626", style: "dotted", width: 1.25 },
};
line.series.items[1].errorBars = {
  type: "plus",
  valueType: "custom",
  plusFormula: "'Summary'!C2:C4",
  plusValues: [5, 7, 9],
  plusFormatCode: "0.0",
  line: { fill: "#EA580C", width: 1 },
};
line.setPosition("M1", "Q10");
assert.match(line.toSvg(), /data-trendline-type="linear"/);
assert.match(line.toSvg(), /data-trendline-type="movingAvg"/);
assert.match(line.toSvg(), /data-trendline-type="poly"/);
assert.match(line.toSvg(), /data-error-bars-series="0"/);
assert.match(line.toSvg(), /data-error-bars-series="1"/);
const pie = sheet.charts.add("pie", sheet.getRange("A1:B4"));
pie.name = "Pie chart";
pie.title = "Revenue share";
pie.setPosition("H12", "L22");
const area = sheet.charts.add("area", sheet.getRange("A1:C4"));
area.name = "Area chart";
area.title = "Revenue and cost area";
area.setPosition("M12", "Q22");
const doughnut = sheet.charts.add("doughnut", sheet.getRange("A1:B4"));
doughnut.name = "Doughnut chart";
doughnut.title = "Revenue mix";
doughnut.dataLabels = { showCategoryName: true, showPercent: true, position: "outsideEnd" };
doughnut.setPosition("H24", "L34");
sheet.getRange("S1:T4").values = [["Units", "Price"], [10, 30], [20, 55], [30, 88]];
const scatter = sheet.charts.add("scatter", sheet.getRange("S1:T4"));
scatter.name = "Scatter chart";
scatter.title = "Price relationship";
scatter.xAxis = { title: { text: "Units" }, min: 0, max: 40, majorUnit: 10, numberFormatCode: "0" };
scatter.yAxis = { title: { text: "Price" }, min: 0, max: 100, majorUnit: 20, numberFormatCode: "$0" };
scatter.series.items[0].marker = { symbol: "diamond", size: 8, fill: "#38BDF8" };
scatter.setPosition("M24", "Q34");
const statusImage = sheet.images.add({
  name: "Status mark",
  alt: "Green status marker",
  accessibility: { title: "Overall status", description: "Green status marker" },
  dataUrl: PNG_DATA_URL,
  anchor: { from: { row: 6, col: 0, rowOffsetPx: 4, colOffsetPx: 4 }, extent: { widthPx: 64, heightPx: 48 } },
});
assert.deepEqual(statusImage.accessibilityCapability, { sourceBound: false, editable: true, addable: true });

const accessibilityWorkbook = Workbook.create();
const accessibilitySheet = accessibilityWorkbook.worksheets.add("Audit");
accessibilitySheet.getRange("A1:B2").values = [["Category", "Value"], ["Ready", 1]];
const unnamedAltImage = accessibilitySheet.images.add({ name: "Internal filename.png", dataUrl: PNG_DATA_URL });
const accessibilityChart = accessibilitySheet.charts.add("bar", accessibilitySheet.getRange("A1:B2"));
assert.equal(unnamedAltImage.alt, "");
const incompleteAccessibilityAudit = accessibilityWorkbook.auditAccessibility({ maxChars: 20_000 });
assert.equal(incompleteAccessibilityAudit.conformanceClaimed, false);
assert.equal(incompleteAccessibilityAudit.machineCheckPassed, false);
assert.deepEqual(incompleteAccessibilityAudit.issues.map((entry) => entry.type), ["unclassifiedDrawing", "unclassifiedDrawing"]);
assert.deepEqual(incompleteAccessibilityAudit.issues.map((entry) => entry.objectKind), ["image", "chart"]);
assert.throws(() => accessibilityWorkbook.auditAccessibility([]), /options must be an object/i);
assert.throws(() => unnamedAltImage.setAccessibilityMetadata({ decorative: true, description: "conflict" }), /cannot combine decorative/i);
unnamedAltImage.setAccessibilityMetadata({ description: "A green readiness indicator." });
accessibilityChart.setAccessibilityMetadata({ decorative: true });
const completeAccessibilityAudit = accessibilityWorkbook.auditAccessibility();
assert.equal(completeAccessibilityAudit.machineCheckPassed, true);
assert.equal(completeAccessibilityAudit.manualReviewRequired, true);
assert.deepEqual(completeAccessibilityAudit.summary, { sheets: 1, drawings: 2, meaningfulDrawings: 1, decorativeDrawings: 1, unclassifiedDrawings: 0, missingTextDrawings: 0 });
unnamedAltImage.setAccessibilityMetadata({ decorative: true, description: null });
const decorativeImageSvg = unnamedAltImage.toSvg();
assert.match(decorativeImageSvg, /^<g aria-hidden="true"><image\b/);
assert.doesNotMatch(decorativeImageSvg, /<(?:title|desc)>/);

workbook.comments.setSelf({ displayName: "Spreadsheet Agent" });
const marginReviewThread = workbook.comments.addThread(
  { cell: sheet.getRange("F2") },
  "Check the calculated margin.",
  {
    id: "margin-review",
    author: "Reviewer",
    resolved: true,
    comment: {
      id: "{11111111-1111-4111-8111-111111111111}",
      personId: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
      date: "2026-07-16T09:00:00.000Z",
      person: {
        id: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
        displayName: "Reviewer",
        userId: "reviewer@example.com",
        providerId: "None",
      },
    },
  },
);
marginReviewThread.addReply("Confirmed against the source workbook.", {
  id: "{22222222-2222-4222-8222-222222222222}",
  personId: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
  author: "Lead reviewer",
  date: "2026-07-16T09:30:00.000Z",
  person: {
    id: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
    displayName: "Lead reviewer",
    userId: "lead@example.com",
    providerId: "None",
  },
  done: true,
});

workbook.recalculate();
assert.deepEqual(sheet.getRange("F2:F4").values, [[0.4], [0.4166666666666667], [0.4]]);
assert.equal(sheet.getRange("E2").values[0][0] instanceof Date, true);
const modelInspect = workbook.inspect({
  kind: "workbook,sheet,worksheetProtection,table,formula,style,computedStyle,drawing,dataValidation,conditionalFormat,thread",
  sheetName: "Summary",
  range: "A1:Q34",
  maxChars: 32_000,
});
assert.match(modelInspect.ndjson, /"name":"SummaryTable"/);
assert.match(modelInspect.ndjson, /"formula":"=\(B2-C2\)\/B2"/);
assert.match(modelInspect.ndjson, /"drawingType":"chart"/);
assert.match(modelInspect.ndjson, /"kind":"dataValidation"/);
assert.match(modelInspect.ndjson, /"kind":"conditionalFormat"/);
assert.match(modelInspect.ndjson, /"kind":"worksheetProtection"/);
assert.match(modelInspect.ndjson, /"kind":"dataBar"/);
assert.match(modelInspect.ndjson, /"kind":"iconSet"/);
assert.match(modelInspect.ndjson, /Check the calculated margin/);
const modelLayout = sheet.layoutJson({ range: "B2:C4" });
assert.equal(modelLayout.cells.find((cell) => cell.address === "B2").conditionalFormats.find((rule) => rule.ruleType === "dataBar").visual.showValue, false);
assert.equal(modelLayout.cells.find((cell) => cell.address === "C4").conditionalFormats.find((rule) => rule.ruleType === "iconSet").visual.reverse, true);
const modelSvg = sheet.toSvg();
assert.match(modelSvg, /id="data-bar-1-1-0"/);
assert.match(modelSvg, /[▼➜▲]/);
assert.match(modelSvg, /<g role="img"><title>Overall status<\/title><desc>Green status marker<\/desc><image\b/);
assert.doesNotMatch(modelSvg, /<text[^>]*>Green status marker<\/text>/);
const modelVerification = verifyArtifact(workbook);
assert.equal(modelVerification.ok, true, modelVerification.ndjson);

const firstXlsx = await SpreadsheetFile.exportXlsx(workbook);
assert.equal(firstXlsx.metadata.codec, "office-kit");
const packageInspect = await SpreadsheetFile.inspectXlsx(firstXlsx, { maxChars: 32_000 });
assert.equal(packageInspect.ok, true, packageInspect.ndjson);
assert.equal(packageInspect.records[0].semanticIssues, 0);

const firstZip = await JSZip.loadAsync(new Uint8Array(await firstXlsx.arrayBuffer()));
assert.ok(firstZip.file("xl/workbook.xml"));
assert.ok(firstZip.file("xl/worksheets/sheet1.xml"));
assert.ok(firstZip.file("xl/tables/table1.xml"));
assert.ok(firstZip.file("xl/drawings/drawing1.xml"));
const firstChartPaths = Object.keys(firstZip.files).filter((name) => /\/charts\/chart\d+\.xml$/i.test(name));
assert.equal(firstChartPaths.length, 6);
const firstChartXml = await Promise.all(firstChartPaths.map((name) => firstZip.file(name).async("text")));
assert.equal(firstChartXml.filter((xml) => /<c:scatterChart>/.test(xml)).length, 1);
assert.match(firstChartXml.find((xml) => /<c:scatterChart>/.test(xml)), /<c:xVal>[\s\S]*<c:yVal>/);
assert.match(firstChartXml.find((xml) => /<c:doughnutChart>/.test(xml)), /<c:showPercent val="1"\s*\/>/);
const firstLineChartIndex = firstChartXml.findIndex((xml) => /<c:lineChart>/.test(xml));
assert.notEqual(firstLineChartIndex, -1);
const firstLineChartXml = firstChartXml[firstLineChartIndex];
assert.equal((firstLineChartXml.match(/<c:trendline>/g) || []).length, 3);
assert.deepEqual([...firstLineChartXml.matchAll(/<c:trendlineType val="([^"]+)"\s*\/>/g)].map((match) => match[1]), ["linear", "movingAvg", "poly"]);
assert.match(firstLineChartXml, /<c:forward val="0\.5"\s*\/>/);
assert.match(firstLineChartXml, /<c:dispRSqr val="1"\s*\/>/);
assert.equal((firstLineChartXml.match(/<c:errBars>/g) || []).length, 2);
assert.match(firstLineChartXml, /<c:errValType val="percentage"\s*\/>[\s\S]*?<c:val val="10"\s*\/>/);
assert.match(firstLineChartXml, /<c:errBarType val="plus"\s*\/>[\s\S]*?<c:numRef><c:f>'Summary'!C2:C4<\/c:f>/);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/media\//i.test(name)).length, 1);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name)).length, 1);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/persons\/[^/]+\.xml$/i.test(name)).length, 1);
assert.equal(firstZip.file("customXml/office-kit-artifact.json"), null);
const firstWorksheetXml = await firstZip.file("xl/worksheets/sheet1.xml").async("text");
const firstProtectionXml = firstWorksheetXml.match(/<x:sheetProtection\b[^>]*\/>/)?.[0] || "";
assert.match(firstProtectionXml, /sheet="1"/);
assert.match(firstProtectionXml, /selectLockedCells="0"/);
assert.match(firstProtectionXml, /selectUnlockedCells="0"/);
assert.match(firstProtectionXml, /sort="0"/);
assert.match(firstProtectionXml, /autoFilter="0"/);
assert.match(firstProtectionXml, /formatCells="1"/);
assert.doesNotMatch(firstProtectionXml, /password=|algorithmName=|hashValue=|saltValue=|spinCount=/);
assert.match(firstWorksheetXml, /<x:mergeCell ref="A6:F6"/);
assert.match(firstWorksheetXml, /<x:dataValidations count="3">/);
assert.match(firstWorksheetXml, /type="list" errorStyle="warning" allowBlank="0" showDropDown="0" showInputMessage="1" showErrorMessage="1" errorTitle="Invalid status" error="Choose a value from the list\." promptTitle="Choose a status" prompt="Use one approved workflow state\." sqref="D2:D4"/);
assert.match(firstWorksheetXml, /type="custom" errorStyle="stop" allowBlank="1" showErrorMessage="1" errorTitle="Cost exceeds revenue" error="Enter a cost no greater than revenue\." sqref="C2:C4"/);
assert.equal((firstWorksheetXml.match(/<x:conditionalFormatting\b/g) || []).length, 6);
assert.match(firstWorksheetXml, /<x:dataBar showValue="0">[\s\S]*?<x:cfvo type="min"\s*\/>[\s\S]*?<x:cfvo type="max"\s*\/>[\s\S]*?<x:color rgb="FF2563EB"\s*\/>/);
assert.match(firstWorksheetXml, /<x:iconSet iconSet="3Arrows" reverse="1">[\s\S]*?<x:cfvo type="num" val="0"\s*\/>[\s\S]*?<x:cfvo type="percent" val="50"\s*\/>[\s\S]*?<x:cfvo type="percent" val="80"\s*\/>/);
const firstThreadedPart = Object.keys(firstZip.files).find((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name));
const firstThreadedXml = await firstZip.file(firstThreadedPart).async("text");
assert.match(firstThreadedXml, /parentId="\{11111111-1111-4111-8111-111111111111\}"/);
assert.match(firstThreadedXml, /Confirmed against the source workbook\./);

const imported = await SpreadsheetFile.importXlsx(firstXlsx);
const importedSheet = imported.worksheets.getItem("Summary");
assert.ok(importedSheet);
assert.deepEqual(importedSheet.getRange("F2:F4").values, [[0.4], [0.4166666666666667], [0.4]]);
assert.deepEqual(importedSheet.getRange("F2:F4").formulas, [
  ["=(B2-C2)/B2"],
  ["=(B3-C3)/B3"],
  ["=(B4-C4)/B4"],
]);
assert.equal(typeof importedSheet.getRange("E2").values[0][0], "number");
assert.equal(importedSheet.getRange("E2").format.numberFormat, "yyyy-mm-dd");
assert.equal(importedSheet.getRange("A1").format.fill, "#0F766E");
assert.equal(importedSheet.getRange("A1").format.font.bold, true);
assert.equal(importedSheet.getRange("A1").format.border.bottom.style, "double");
assert.deepEqual(importedSheet.getRange("F2").format.protection, { locked: false, hidden: true });
assert.deepEqual(importedSheet.mergedRanges, ["A6:F6"]);
assert.ok(Math.abs(importedSheet.getRange("A1:A6").format.columnWidthPx - 96) <= 1);
assert.ok(Math.abs(importedSheet.getRange("A1:F1").format.rowHeightPx - 28) < 0.01);
assert.equal(importedSheet.getRange("A8:F8").format.rowHidden, true);
assert.deepEqual(importedSheet.freezePanes.toJSON(), { rows: 1, columns: 1, frozen: true, topLeftCell: "B2", activePane: "bottomRight" });
assert.deepEqual(importedSheet.protection, {
  enabled: true,
  allow: ["selectLockedCells", "selectUnlockedCells", "sort", "autoFilter"],
});
assert.equal(importedSheet.tables.items[0].name, "SummaryTable");
assert.equal(importedSheet.images.items[0].alt, "Green status marker");
assert.deepEqual(importedSheet.images.items[0].accessibility, { title: "Overall status", description: "Green status marker" });
assert.deepEqual(importedSheet.images.items[0].accessibilityCapability, { sourceBound: true, editable: true, addable: true });
assert.deepEqual(importedSheet.charts.items.map((chart) => chart.type), ["bar", "line", "pie", "area", "doughnut", "scatter"]);
assert.deepEqual(importedSheet.charts.items[0].accessibility, { title: "Quarterly revenue and cost chart", description: "Revenue and cost both rise from January through March." });
assert.deepEqual(importedSheet.charts.items[0].accessibilityCapability, { sourceBound: true, editable: true, addable: true });
const importedLine = importedSheet.charts.items[1];
assert.deepEqual(importedLine.series.items[0].trendlines.map((trendline) => trendline.type), ["linear", "movingAvg", "poly"]);
assert.equal(importedLine.series.items[0].trendlines[0].name, "Revenue projection");
assert.equal(importedLine.series.items[0].trendlines[0].forward, 0.5);
assert.equal(importedLine.series.items[0].trendlines[0].displayEquation, true);
assert.equal(importedLine.series.items[0].trendlines[0].displayRSquared, true);
assert.deepEqual(importedLine.series.items[0].trendlines[0].line, { fill: "#7C3AED", style: "dashed", width: 1.5 });
assert.match(importedLine.toSvg(), /data-trendline-type="linear"/);
assert.deepEqual(importedLine.series.items[0].errorBars, {
  direction: "y",
  type: "both",
  valueType: "percentage",
  value: 10,
  noEndCap: true,
  line: { fill: "#DC2626", style: "dotted", width: 1.25 },
});
assert.deepEqual(importedLine.series.items[1].errorBars, {
  direction: "y",
  type: "plus",
  valueType: "cust",
  plusValues: [5, 7, 9],
  plusFormula: "'Summary'!C2:C4",
  plusFormatCode: "0.0",
  noEndCap: false,
  line: { fill: "#EA580C", width: 1 },
});
assert.match(importedLine.toSvg(), /data-error-bars-series="0"/);
assert.match(importedSheet.charts.items[3].toSvg(), /data-series-index="0"/);
assert.match(importedSheet.charts.items[4].toSvg(), /data-point-index="0"/);
assert.equal(importedSheet.charts.items[4].dataLabels.showPercent, true);
assert.match(importedSheet.charts.items[4].toSvg(), /data-chart-label-index="0"[^>]*>[^<]*%/);
const importedScatter = importedSheet.charts.items[5];
assert.deepEqual(importedScatter.categories, []);
assert.deepEqual(importedScatter.series.items[0].xValues, [10, 20, 30]);
assert.deepEqual(importedScatter.series.items[0].values, [30, 55, 88]);
assert.equal(importedScatter.series.items[0].xFormula, "'Summary'!S2:S4");
assert.equal(importedScatter.xAxis.axisType, "valueAxis");
assert.match(importedScatter.toSvg(), /<polygon[^>]+38BDF8/i);
assert.deepEqual(importedSheet.dataValidations.items.map((item) => item.rule.type), ["list", "whole", "custom"]);
const importedListValidation = importedSheet.dataValidations.items.find((item) => item.rule.type === "list");
assert.deepEqual(importedListValidation.rule, {
  type: "list",
  values: ["Planned", "Review", "Done"],
  allowBlank: false,
  showInputMessage: true,
  promptTitle: "Choose a status",
  prompt: "Use one approved workflow state.",
  showErrorMessage: true,
  errorTitle: "Invalid status",
  error: "Choose a value from the list.",
  errorStyle: "warning",
  showDropdown: true,
});
assert.deepEqual(importedSheet.dataValidations.items.find((item) => item.rule.type === "custom").rule, {
  type: "custom",
  formula1: "=C2<=B2",
  allowBlank: true,
  showErrorMessage: true,
  errorTitle: "Cost exceeds revenue",
  error: "Enter a cost no greater than revenue.",
  errorStyle: "stop",
});
assert.deepEqual(importedSheet.conditionalFormattings.items.map((item) => item.ruleType), ["cellIs", "expression", "containsText", "colorScale", "dataBar", "iconSet"]);
const importedDataBar = importedSheet.conditionalFormattings.items.find((item) => item.ruleType === "dataBar");
assert.equal(importedDataBar.color, "#2563EB");
assert.equal(importedDataBar.showValue, false);
assert.deepEqual(importedDataBar.thresholds, [{ type: "min" }, { type: "max" }]);
const importedIconSet = importedSheet.conditionalFormattings.items.find((item) => item.ruleType === "iconSet");
assert.equal(importedIconSet.iconSet, "3Arrows");
assert.equal(importedIconSet.reverse, true);
assert.deepEqual(importedIconSet.thresholds, [{ type: "num", value: 0 }, { type: "percent", value: 50 }, { type: "percent", value: 80 }]);
assert.equal(imported.comments.threads.length, 1);
assert.equal(imported.comments.threads[0].comments.length, 2);
assert.equal(imported.comments.threads[0].comments[0].text, "Check the calculated margin.");
assert.equal(imported.comments.threads[0].comments[1].text, "Confirmed against the source workbook.");
assert.equal(imported.comments.threads[0].resolved, true);

importedSheet.getRange("B2").values = [[110]];
importedSheet.getRange("E2").values = [[new Date("2026-01-20T00:00:00.000Z")]];
importedSheet.getRange("F2").formulas = [["=(B2-C2)/B2"]];
importedSheet.getRange("F2").format.fill = "#BBF7D0";
importedSheet.getRange("A1:A6").format.columnWidthPx = 104;
importedSheet.getRange("A1:F1").format.rowHeightPx = 30;
importedSheet.freezePanes.unfreeze();
importedSheet.freezePanes.freezeRows(2);
importedSheet.freezePanes.freezeColumns(1);
importedSheet.protection = { allow: ["selectUnlockedCells", "formatCells"] };
importedSheet.tables.items[0].style = "TableStyleMedium9";
importedSheet.images.items[0].alt = "Edited green status marker";
importedSheet.charts.items[1].title = "Edited revenue trend";
importedLine.series.items[0].trendlines[0].name = "Edited revenue projection";
importedLine.series.items[0].trendlines[0].forward = 1.5;
importedLine.series.items[0].trendlines[0].line.fill = "#0EA5E9";
importedLine.series.items[0].errorBars.value = 15;
importedLine.series.items[0].errorBars.line.fill = "#BE123C";
importedLine.series.items[1].errorBars.plusValues[1] = 8;
importedScatter.title = "Edited price relationship";
importedScatter.series.items[0].xValues[1] = 22;
importedScatter.series.items[0].values[1] = 60;
const listValidation = importedSheet.dataValidations.items.find((item) => item.rule.type === "list");
listValidation.rule.values.push("Blocked");
listValidation.rule.prompt = "Pick the current workflow state.";
listValidation.rule.errorStyle = "information";
listValidation.rule.showDropdown = false;
const marginConditional = importedSheet.conditionalFormattings.items.find((item) => item.ruleType === "cellIs");
marginConditional.formula = "0.45";
marginConditional.format.fill = "#BBF7D0";
importedDataBar.color = "#0EA5E9";
importedDataBar.showValue = true;
importedIconSet.thresholds[1].value = 60;
importedIconSet.reverse = false;
const importedThread = imported.comments.threads[0];
importedThread.comments[0].text = "Margin reviewed after edit.";
importedThread.comments[1].text = "Reply reviewed after edit.";
importedThread.reopen();
imported.recalculate();
assert.equal(importedSheet.getRange("F2").values[0][0], 50 / 110);

const secondXlsx = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
const second = await SpreadsheetFile.importXlsx(secondXlsx);
const secondSheet = second.worksheets.getItem("Summary");
assert.equal(secondSheet.getRange("B2").values[0][0], 110);
assert.equal(secondSheet.getRange("F2").values[0][0], 50 / 110);
assert.equal(secondSheet.getRange("F2").format.fill, "#BBF7D0");
assert.ok(Math.abs(secondSheet.getRange("A1:A6").format.columnWidthPx - 104) <= 1);
assert.ok(Math.abs(secondSheet.getRange("A1:F1").format.rowHeightPx - 30) < 0.01);
assert.deepEqual(secondSheet.freezePanes.toJSON(), { rows: 2, columns: 1, frozen: true, topLeftCell: "B3", activePane: "bottomRight" });
assert.deepEqual(secondSheet.protection, { enabled: true, allow: ["selectUnlockedCells", "formatCells"] });
assert.equal(secondSheet.tables.items[0].style, "TableStyleMedium9");
assert.equal(secondSheet.images.items[0].alt, "Edited green status marker");
assert.equal(secondSheet.charts.items[1].title, "Edited revenue trend");
assert.equal(secondSheet.charts.items[1].series.items[0].trendlines[0].name, "Edited revenue projection");
assert.equal(secondSheet.charts.items[1].series.items[0].trendlines[0].forward, 1.5);
assert.equal(secondSheet.charts.items[1].series.items[0].trendlines[0].line.fill, "#0EA5E9");
assert.equal(secondSheet.charts.items[1].series.items[0].errorBars.value, 15);
assert.equal(secondSheet.charts.items[1].series.items[0].errorBars.line.fill, "#BE123C");
assert.deepEqual(secondSheet.charts.items[1].series.items[1].errorBars.plusValues, [5, 8, 9]);
assert.equal(secondSheet.charts.items[1].series.items[1].errorBars.plusFormula, "'Summary'!C2:C4");
assert.equal(secondSheet.charts.items[4].dataLabels.showPercent, true);
assert.equal(secondSheet.charts.items[5].title, "Edited price relationship");
assert.deepEqual(secondSheet.charts.items[5].series.items[0].xValues, [10, 22, 30]);
assert.deepEqual(secondSheet.charts.items[5].series.items[0].values, [30, 60, 88]);
assert.deepEqual(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.values, ["Planned", "Review", "Done", "Blocked"]);
assert.equal(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.prompt, "Pick the current workflow state.");
assert.equal(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.errorStyle, "information");
assert.equal(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.showDropdown, false);
assert.equal(secondSheet.dataValidations.items.find((item) => item.rule.type === "custom").rule.formula1, "=C2<=B2");
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "cellIs").formula, "0.45");
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "dataBar").color, "#0EA5E9");
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "dataBar").showValue, true);
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "iconSet").thresholds[1].value, 60);
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "iconSet").reverse, false);
assert.equal(second.comments.threads[0].comments.length, 2);
assert.equal(second.comments.threads[0].comments[0].text, "Margin reviewed after edit.");
assert.equal(second.comments.threads[0].comments[1].text, "Reply reviewed after edit.");
assert.equal(second.comments.threads[0].resolved, false);

secondSheet.protection = null;
const protectionRemoved = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(second, { recalculate: false }));
assert.equal(protectionRemoved.worksheets.getItem("Summary").protection, undefined);

const secondInspect = second.inspect({ kind: "workbook,sheet,table,formula,style,drawing,dataValidation,conditionalFormat,thread", maxChars: 32_000 });
assert.match(secondInspect.ndjson, /Edited revenue trend/);
assert.match(secondInspect.ndjson, /Margin reviewed after edit/);
const secondVerification = verifyArtifact(second);
assert.equal(secondVerification.ok, true, secondVerification.ndjson);
const secondPackageInspect = await SpreadsheetFile.inspectXlsx(secondXlsx, { maxChars: 32_000 });
assert.equal(secondPackageInspect.ok, true, secondPackageInspect.ndjson);
assert.equal(secondPackageInspect.records[0].semanticIssues, 0);

const bubbleWorkbook = Workbook.create();
const bubbleSheet = bubbleWorkbook.worksheets.add("Opportunities");
bubbleSheet.getRange("A1:C4").values = [
  ["Customers", "Revenue", "Pipeline"],
  [10, 42, 4],
  [20, 68, 9],
  [30, 85, 16],
];
const bubble = bubbleSheet.charts.add("bubble", bubbleSheet.getRange("A1:C4"));
bubble.name = "Opportunity bubble";
bubble.title = "Revenue opportunity";
bubble.xAxis = { title: { text: "Customers" }, min: 0, max: 40, majorUnit: 10, numberFormatCode: "0" };
bubble.yAxis = { title: { text: "Revenue" }, min: 0, max: 100, majorUnit: 20, numberFormatCode: "$0" };
bubble.series.items[0].fill = "#0EA5E9";
bubble.series.items[0].line = { fill: "#0369A1", width: 1.5 };
bubble.setPosition("E2", "L18");
assert.deepEqual(bubble.categories, []);
assert.deepEqual(bubble.series.items[0].xValues, [10, 20, 30]);
assert.deepEqual(bubble.series.items[0].values, [42, 68, 85]);
assert.deepEqual(bubble.series.items[0].bubbleSizes, [4, 9, 16]);
assert.equal(bubble.series.items[0].xFormula, "'Opportunities'!A2:A4");
assert.equal(bubble.series.items[0].formula, "'Opportunities'!B2:B4");
assert.equal(bubble.series.items[0].bubbleSizeFormula, "'Opportunities'!C2:C4");
assert.match(bubble.toSvg(), /data-bubble-size="16"/);
assert.equal(bubbleWorkbook.verify().ok, true);
const bubbleXlsx = await SpreadsheetFile.exportXlsx(bubbleWorkbook);
const bubbleZip = await JSZip.loadAsync(new Uint8Array(await bubbleXlsx.arrayBuffer()));
const bubbleChartPath = Object.keys(bubbleZip.files).find((name) => /\/charts\/chart\d+\.xml$/i.test(name));
assert.ok(bubbleChartPath);
const bubbleXml = await bubbleZip.file(bubbleChartPath).async("text");
assert.match(bubbleXml, /<c:bubbleChart>/);
assert.match(bubbleXml, /<c:xVal>[\s\S]*<c:yVal>[\s\S]*<c:bubbleSize>/);
assert.equal((bubbleXml.match(/<c:valAx>/g) || []).length, 2);
assert.doesNotMatch(bubbleXml, /<c:cat>/);
const importedBubbleWorkbook = await SpreadsheetFile.importXlsx(bubbleXlsx);
const importedBubble = importedBubbleWorkbook.worksheets.getItem("Opportunities").charts.items[0];
assert.equal(importedBubble.type, "bubble");
assert.equal(importedBubble.xAxis.axisType, "valueAxis");
assert.deepEqual(importedBubble.series.items[0].bubbleSizes, [4, 9, 16]);
importedBubble.title = "Edited revenue opportunity";
importedBubble.series.items[0].xValues[1] = 22;
importedBubble.series.items[0].values[1] = 70;
importedBubble.series.items[0].bubbleSizes[1] = 12;
const editedBubbleWorkbook = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(importedBubbleWorkbook));
const editedBubble = editedBubbleWorkbook.worksheets.getItem("Opportunities").charts.items[0];
assert.equal(editedBubble.title, "Edited revenue opportunity");
assert.deepEqual(editedBubble.series.items[0].xValues, [10, 22, 30]);
assert.deepEqual(editedBubble.series.items[0].values, [42, 70, 85]);
assert.deepEqual(editedBubble.series.items[0].bubbleSizes, [4, 12, 16]);

function chartBoundaryWorkbook(type) {
  const candidate = Workbook.create();
  const candidateSheet = candidate.worksheets.add("Chart boundary");
  candidateSheet.getRange("A1:B3").values = [["Quarter", "Revenue"], ["Q1", 40], ["Q2", 60]];
  return { candidate, chart: candidateSheet.charts.add(type, candidateSheet.getRange("A1:B3")) };
}

const xDirectionPreview = chartBoundaryWorkbook("line").chart;
const xDirectionBasePoints = /<polyline points="([^"]+)"[^>]*data-series-index="0"/.exec(xDirectionPreview.toSvg())?.[1];
xDirectionPreview.series.items[0].errorBars = { direction: "x", valueType: "fixedVal", value: 0.25 };
const xDirectionSvg = xDirectionPreview.toSvg();
assert.equal(/<polyline points="([^"]+)"[^>]*data-series-index="0"/.exec(xDirectionSvg)?.[1], xDirectionBasePoints);
const xDirectionMark = /<line data-error-bars-series="0" data-error-bars-index="0" x1="([^"]+)" y1="([^"]+)" x2="([^"]+)" y2="([^"]+)"/.exec(xDirectionSvg);
assert.ok(xDirectionMark);
assert.notEqual(xDirectionMark[1], xDirectionMark[3]);
assert.equal(xDirectionMark[2], xDirectionMark[4]);

const invalidDoughnutAxes = chartBoundaryWorkbook("doughnut");
invalidDoughnutAxes.chart.xAxis = {};
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidDoughnutAxes.candidate),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /doughnut charts cannot carry.*axes/i.test(error.message),
);

const invalidAreaLineOptions = chartBoundaryWorkbook("area");
invalidAreaLineOptions.chart.lineOptions = { grouping: "standard" };
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidAreaLineOptions.candidate),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /lineOptions require a line chart/i.test(error.message),
);

const invalidAreaTrendline = chartBoundaryWorkbook("area");
invalidAreaTrendline.chart.series.items[0].trendlines = [{ type: "linear" }];
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidAreaTrendline.candidate),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /supported only for bar and line/i.test(error.message),
);

const invalidAreaErrorBars = chartBoundaryWorkbook("area");
invalidAreaErrorBars.chart.series.items[0].errorBars = { type: "percentage", value: 5 };
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidAreaErrorBars.candidate),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /supported only for bar and line/i.test(error.message),
);

const invalidCustomErrorBars = chartBoundaryWorkbook("line");
invalidCustomErrorBars.chart.series.items[0].errorBars = { valueType: "custom", plusValues: [1, 2] };
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidCustomErrorBars.candidate),
  (error) => error?.code === "invalid_spreadsheet_chart" && /minus requires literal values or a formula/i.test(error.message),
);

const unknownErrorBarField = chartBoundaryWorkbook("line");
unknownErrorBarField.chart.series.items[0].errorBars = { type: "percentage", value: 5, confidence: 0.95 };
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(unknownErrorBarField.candidate),
  (error) => error?.code === "invalid_spreadsheet_chart" && /unsupported fields: confidence/i.test(error.message),
);

const invalidMovingAverage = chartBoundaryWorkbook("line");
invalidMovingAverage.chart.series.items[0].trendlines = [{ type: "movingAvg", period: 2 }];
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidMovingAverage.candidate),
  (error) => error?.code === "invalid_spreadsheet_chart" && /require at least three series values/i.test(error.message),
);

const invalidPolynomial = chartBoundaryWorkbook("line");
invalidPolynomial.chart.series.items[0].trendlines = [{ type: "poly", order: 7 }];
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidPolynomial.candidate),
  (error) => error?.code === "invalid_spreadsheet_chart" && /order must be an integer from 2 to 6/i.test(error.message),
);

const changedTrendlineTopology = await SpreadsheetFile.importXlsx(firstXlsx);
changedTrendlineTopology.worksheets.getItem("Summary").charts.items[1].series.items[0].trendlines.pop();
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(changedTrendlineTopology),
  (error) => error?.code === "unsupported_spreadsheet_chart_edit" && /cannot change imported trendline count/i.test(error.message),
);

const changedErrorBarTopology = await SpreadsheetFile.importXlsx(firstXlsx);
changedErrorBarTopology.worksheets.getItem("Summary").charts.items[1].series.items[0].errorBars = undefined;
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(changedErrorBarTopology),
  (error) => error?.code === "unsupported_spreadsheet_chart_edit" && /cannot add or remove imported error bars/i.test(error.message),
);

const unsupportedTrendlineZip = await JSZip.loadAsync(new Uint8Array(await firstXlsx.arrayBuffer()));
const unsupportedTrendlinePath = firstChartPaths[firstLineChartIndex];
const unsupportedTrendlineXml = firstLineChartXml.replace("</c:trendline>", "<c:trendlineLbl/></c:trendline>");
assert.notEqual(unsupportedTrendlineXml, firstLineChartXml);
unsupportedTrendlineZip.file(unsupportedTrendlinePath, unsupportedTrendlineXml);
const unsupportedTrendlineSource = new FileBlob(
  await unsupportedTrendlineZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
  { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", name: "unsupported-trendline-label.xlsx" },
);
const preservedUnsupportedTrendline = await SpreadsheetFile.importXlsx(unsupportedTrendlineSource);
const preservedUnsupportedChart = preservedUnsupportedTrendline.worksheets.getItem("Summary").charts.items[1];
assert.equal(preservedUnsupportedChart.series.items[0].trendlines, undefined);
const preservedUnsupportedOutput = await SpreadsheetFile.exportXlsx(preservedUnsupportedTrendline);
const preservedUnsupportedZip = await JSZip.loadAsync(new Uint8Array(await preservedUnsupportedOutput.arrayBuffer()));
assert.equal(await preservedUnsupportedZip.file(unsupportedTrendlinePath).async("text"), unsupportedTrendlineXml);
assert.deepEqual(preservedUnsupportedChart.accessibilityCapability, { sourceBound: true, editable: true, addable: true });
preservedUnsupportedChart.name = "Renamed residual chart";
preservedUnsupportedChart.setAccessibilityMetadata({ description: "A line chart with a source-owned trendline label." });
const accessibleUnsupportedOutput = await SpreadsheetFile.exportXlsx(preservedUnsupportedTrendline);
const accessibleUnsupportedZip = await JSZip.loadAsync(new Uint8Array(await accessibleUnsupportedOutput.arrayBuffer()));
assert.equal(await accessibleUnsupportedZip.file(unsupportedTrendlinePath).async("text"), unsupportedTrendlineXml);
assert.match(await accessibleUnsupportedZip.file("xl/drawings/drawing1.xml").async("text"), /name="Renamed residual chart"[^>]*descr="A line chart with a source-owned trendline label\."/);
preservedUnsupportedChart.title = "Forbidden trendline-label edit";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(preservedUnsupportedTrendline),
  (error) => error?.code === "unsupported_spreadsheet_chart_edit" && /read-only/i.test(error.message),
);

const unsupportedErrorBarsZip = await JSZip.loadAsync(new Uint8Array(await firstXlsx.arrayBuffer()));
const unsupportedErrorBarsXml = firstLineChartXml.replace("</c:errBars>", "<c:extLst/></c:errBars>");
assert.notEqual(unsupportedErrorBarsXml, firstLineChartXml);
unsupportedErrorBarsZip.file(unsupportedTrendlinePath, unsupportedErrorBarsXml);
const unsupportedErrorBarsSource = new FileBlob(
  await unsupportedErrorBarsZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
  { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", name: "unsupported-error-bars.xlsx" },
);
const preservedUnsupportedErrorBars = await SpreadsheetFile.importXlsx(unsupportedErrorBarsSource);
const preservedUnsupportedErrorBarChart = preservedUnsupportedErrorBars.worksheets.getItem("Summary").charts.items[1];
assert.equal(preservedUnsupportedErrorBarChart.series.items[0].errorBars, undefined);
const preservedUnsupportedErrorBarOutput = await SpreadsheetFile.exportXlsx(preservedUnsupportedErrorBars);
const preservedUnsupportedErrorBarZip = await JSZip.loadAsync(new Uint8Array(await preservedUnsupportedErrorBarOutput.arrayBuffer()));
assert.equal(await preservedUnsupportedErrorBarZip.file(unsupportedTrendlinePath).async("text"), unsupportedErrorBarsXml);
preservedUnsupportedErrorBarChart.title = "Forbidden error-bar extension edit";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(preservedUnsupportedErrorBars),
  (error) => error?.code === "unsupported_spreadsheet_chart_edit" && /read-only/i.test(error.message),
);

const invalidScatter = chartBoundaryWorkbook("scatter");
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidScatter.candidate),
  (error) => error?.code === "invalid_spreadsheet_chart" && /xValue.*finite/i.test(error.message),
);

const scatterWithSeriesLine = Workbook.create();
const scatterWithSeriesLineSheet = scatterWithSeriesLine.worksheets.add("Scatter line boundary");
scatterWithSeriesLineSheet.getRange("A1:B3").values = [["X", "Y"], [1, 2], [2, 4]];
const scatterWithSeriesLineChart = scatterWithSeriesLineSheet.charts.add("scatter", scatterWithSeriesLineSheet.getRange("A1:B3"));
scatterWithSeriesLineChart.series.items[0].line = { fill: "#2563EB", width: 2 };
assert.equal(scatterWithSeriesLine.verify().ok, false);
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(scatterWithSeriesLine),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /marker-only scatter.*marker\.line/i.test(error.message),
);

const bubbleShortcutBoundary = Workbook.create();
const bubbleShortcutBoundarySheet = bubbleShortcutBoundary.worksheets.add("Bubble boundary");
bubbleShortcutBoundarySheet.getRange("A1:D3").values = [["X", "Y", "Size", "Unexpected"], [1, 2, 3, 4], [2, 4, 6, 8]];
assert.throws(
  () => bubbleShortcutBoundarySheet.charts.add("bubble", bubbleShortcutBoundarySheet.getRange("A1:D3")),
  /requires exactly three columns ordered X \| Y \| Size/i,
);
const nonPositiveBubble = Workbook.create();
const nonPositiveBubbleSheet = nonPositiveBubble.worksheets.add("Non-positive bubble");
nonPositiveBubbleSheet.getRange("A1:C3").values = [["X", "Y", "Size"], [1, 2, 0], [2, 4, 6]];
assert.throws(
  () => nonPositiveBubbleSheet.charts.add("bubble", nonPositiveBubbleSheet.getRange("A1:C3")),
  /Size value.*finite and positive/i,
);
const mismatchedBubble = Workbook.create();
const mismatchedBubbleSheet = mismatchedBubble.worksheets.add("Mismatched bubble");
const mismatchedBubbleChart = mismatchedBubbleSheet.charts.add("bubble", {
  name: "Mismatched bubble",
  series: [{ name: "Series", xValues: [1, 2], values: [2, 4], bubbleSizes: [5] }],
});
assert.equal(mismatchedBubbleChart.type, "bubble");
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(mismatchedBubble),
  (error) => error?.code === "invalid_spreadsheet_chart" && /bubbleSizes.*y values/i.test(error.message),
);

const pivotWorkbook = Workbook.create();
const pivotData = pivotWorkbook.worksheets.add("Data");
pivotData.getRange("A1:C5").values = [
  ["Region", "Product", "Sales"],
  ["East", "A", 10],
  ["East", "B", 20],
  ["West", "A", 30],
  ["West", "B", 40],
];
const pivotSummary = pivotWorkbook.worksheets.add("Summary");
pivotSummary.getRange("A1:D4").format = { fill: "#ECFEFF" };
const nativePivot = pivotSummary.pivotTables.add({
  name: "Sales by region",
  sourceRange: "Data!A1:C5",
  targetRange: "A1",
  rowFields: ["Region"],
  columnFields: ["Product"],
  valueFields: [{ field: "Sales", summarizeBy: "sum" }],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
assert.deepEqual(nativePivot.computedValues(), [
  ["Region", "A", "B", "Grand Total"],
  ["East", 10, 20, 30],
  ["West", 30, 40, 70],
  ["Grand Total", 40, 60, 100],
]);
const nativePivotXlsx = await SpreadsheetFile.exportXlsx(pivotWorkbook);
const nativePivotZip = await JSZip.loadAsync(new Uint8Array(await nativePivotXlsx.arrayBuffer()));
const nativePivotPart = Object.keys(nativePivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const nativePivotCache = Object.keys(nativePivotZip.files).find((name) => /pivotCache\/pivotCacheDefinition.*\.xml$/i.test(name));
const nativePivotRecords = Object.keys(nativePivotZip.files).find((name) => /pivotCache\/pivotCacheRecords.*\.xml$/i.test(name));
assert.ok(nativePivotPart);
assert.ok(nativePivotCache);
assert.ok(nativePivotRecords);
assert.match(await nativePivotZip.file(nativePivotPart).async("text"), /name="Sales by region"[\s\S]*location ref="A1:D4"[\s\S]*subtotal="sum"/);
assert.match(await nativePivotZip.file(nativePivotCache).async("text"), /worksheetSource ref="A1:C5" sheet="Data"/);
assert.match(await nativePivotZip.file(nativePivotRecords).async("text"), /count="4"/);

const importedPivotWorkbook = await SpreadsheetFile.importXlsx(nativePivotXlsx);
const importedPivot = importedPivotWorkbook.worksheets.getItem("Summary").pivotTables.items[0];
assert.equal(importedPivot.name, "Sales by region");
assert.deepEqual(importedPivot.rowFields, ["Region"]);
assert.deepEqual(importedPivot.columnFields, ["Product"]);
assert.deepEqual(importedPivot.valueFields, [{ field: "Sales", summarizeBy: "sum", name: "Sum of Sales" }]);
assert.deepEqual(importedPivot.computedValues(), nativePivot.computedValues());
assert.equal(importedPivotWorkbook.worksheets.getItem("Summary").getRange("A1").format.fill, "#ECFEFF");
const secondPivotXlsx = await SpreadsheetFile.exportXlsx(importedPivotWorkbook);
const secondPivotZip = await JSZip.loadAsync(new Uint8Array(await secondPivotXlsx.arrayBuffer()));
assert.equal(await secondPivotZip.file(nativePivotPart).async("text"), await nativePivotZip.file(nativePivotPart).async("text"));
assert.equal(await secondPivotZip.file(nativePivotCache).async("text"), await nativePivotZip.file(nativePivotCache).async("text"));
assert.equal(await secondPivotZip.file(nativePivotRecords).async("text"), await nativePivotZip.file(nativePivotRecords).async("text"));

assert.deepEqual(importedPivot.sourceCapabilities, { sourceBound: true, refreshOnLoadHardenable: true });
assert.deepEqual(importedPivot.inspectRecord().sourceCapabilities, { sourceBound: true, refreshOnLoadHardenable: true });
const hardeningPivotWorkbook = await SpreadsheetFile.importXlsx(nativePivotXlsx);
const hardeningPivot = hardeningPivotWorkbook.worksheets.getItem("Summary").pivotTables.items[0];
hardeningPivot.disableRefreshOnLoad();
assert.equal(hardeningPivot.refreshPolicy.refreshOnLoad, false);
const hardenedPivotXlsx = await SpreadsheetFile.exportXlsx(hardeningPivotWorkbook);
const hardenedPivotZip = await JSZip.loadAsync(new Uint8Array(await hardenedPivotXlsx.arrayBuffer()));
const nativeCacheXml = await nativePivotZip.file(nativePivotCache).async("text");
const hardenedCacheXml = await hardenedPivotZip.file(nativePivotCache).async("text");
assert.match(nativeCacheXml, /refreshOnLoad="1"/);
assert.match(hardenedCacheXml, /refreshOnLoad="0"/);
assert.equal(
  hardenedCacheXml.replace(/\srefreshOnLoad="(?:1|true|TRUE|0|false|FALSE)"/, ""),
  nativeCacheXml.replace(/\srefreshOnLoad="(?:1|true|TRUE|0|false|FALSE)"/, ""),
);
for (const name of Object.keys(nativePivotZip.files).filter((name) => !nativePivotZip.files[name].dir && name !== nativePivotCache)) {
  assert.deepEqual(
    await hardenedPivotZip.file(name).async("uint8array"),
    await nativePivotZip.file(name).async("uint8array"),
    `only ${nativePivotCache} may change during source-bound PivotTable refresh hardening (${name})`,
  );
}
const hardenedPivotReimport = await SpreadsheetFile.importXlsx(hardenedPivotXlsx);
const hardenedPivotReimported = hardenedPivotReimport.worksheets.getItem("Summary").pivotTables.items[0];
assert.equal(hardenedPivotReimported.refreshPolicy.refreshOnLoad, false);
assert.deepEqual(hardenedPivotReimported.sourceCapabilities, { sourceBound: true, refreshOnLoadHardenable: false });
assert.throws(() => hardenedPivotReimported.disableRefreshOnLoad(), /explicit refreshOnLoad=true/i);

const nonRefreshPivotEdit = await SpreadsheetFile.importXlsx(nativePivotXlsx);
nonRefreshPivotEdit.worksheets.getItem("Summary").pivotTables.items[0].refreshPolicy = {
  ...nonRefreshPivotEdit.worksheets.getItem("Summary").pivotTables.items[0].refreshPolicy,
  enableRefresh: false,
};
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(nonRefreshPivotEdit),
  (error) => error?.code === "unsupported_spreadsheet_pivot_edit" && /only refreshOnLoad/i.test(error.message),
);

const omittedRefreshPivotZip = await JSZip.loadAsync(new Uint8Array(await nativePivotXlsx.arrayBuffer()));
omittedRefreshPivotZip.file(nativePivotCache, nativeCacheXml.replace(/\srefreshOnLoad="1"/, ""));
const omittedRefreshPivot = await SpreadsheetFile.importXlsx(new FileBlob(
  await omittedRefreshPivotZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }),
  { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", name: "pivot-refresh-omitted.xlsx" },
));
const omittedRefreshPivotTable = omittedRefreshPivot.worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(omittedRefreshPivotTable.sourceCapabilities, { sourceBound: true, refreshOnLoadHardenable: false });
assert.throws(() => omittedRefreshPivotTable.disableRefreshOnLoad(), /explicit refreshOnLoad=true/i);

const filteredPivotWorkbook = Workbook.create();
const filteredPivotData = filteredPivotWorkbook.worksheets.add("Data");
filteredPivotData.getRange("A1:C5").values = [
  ["Region", "Product", "Sales"],
  ["East", "A", 10],
  ["East", "B", 20],
  ["West", "A", 30],
  ["West", "B", 40],
];
const filteredPivotSummary = filteredPivotWorkbook.worksheets.add("Summary");
filteredPivotSummary.getRange("A1:C3").format = { fill: "#FFF7ED" };
const filteredPivot = filteredPivotSummary.pivotTables.add({
  name: "Filtered sales",
  sourceRange: "Data!A1:C5",
  targetRange: "A1",
  rowFields: ["Region"],
  columnFields: ["Product"],
  valueFields: [{ field: "Sales", summarizeBy: "sum" }],
  filters: [
    { field: "Region", include: ["East"] },
    { field: "Product", exclude: ["B"] },
  ],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
assert.deepEqual(filteredPivot.computedValues(), [
  ["Region", "A", "Grand Total"],
  ["East", 10, 10],
  ["Grand Total", 10, 10],
]);
const filteredPivotXlsx = await SpreadsheetFile.exportXlsx(filteredPivotWorkbook);
const filteredPivotZip = await JSZip.loadAsync(new Uint8Array(await filteredPivotXlsx.arrayBuffer()));
const filteredPivotPart = Object.keys(filteredPivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const filteredPivotXml = await filteredPivotZip.file(filteredPivotPart).async("text");
assert.match(filteredPivotXml, /location ref="A1:C3"/);
assert.match(filteredPivotXml, /includeNewItemsInFilter="0"[\s\S]*<item x="1" h="1"/);
assert.match(filteredPivotXml, /includeNewItemsInFilter="1"[\s\S]*<item x="1" h="1"/);
const importedFilteredPivotWorkbook = await SpreadsheetFile.importXlsx(filteredPivotXlsx);
const importedFilteredPivot = importedFilteredPivotWorkbook.worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(importedFilteredPivot.filters, [
  { field: "Region", include: ["East"] },
  { field: "Product", exclude: ["B"] },
]);
assert.deepEqual(importedFilteredPivot.computedValues(), filteredPivot.computedValues());
assert.equal(importedFilteredPivotWorkbook.worksheets.getItem("Summary").getRange("C3").format.fill, "#FFF7ED");
const secondFilteredPivotXlsx = await SpreadsheetFile.exportXlsx(importedFilteredPivotWorkbook);
const secondFilteredPivotZip = await JSZip.loadAsync(new Uint8Array(await secondFilteredPivotXlsx.arrayBuffer()));
assert.equal(await secondFilteredPivotZip.file(filteredPivotPart).async("text"), filteredPivotXml);
importedFilteredPivot.filters[0].include[0] = "West";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(importedFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_edit" && /read-only/i.test(error.message),
);

const typedFilteredPivotWorkbook = Workbook.create();
const typedFilteredPivotData = typedFilteredPivotWorkbook.worksheets.add("Data");
typedFilteredPivotData.getRange("A1:B5").values = [["Key", "Sales"], [1, 10], [true, 20], [null, 30], ["Other", 40]];
const typedFilteredPivotSummary = typedFilteredPivotWorkbook.worksheets.add("Summary");
const typedFilteredPivot = typedFilteredPivotSummary.pivotTables.add({
  name: "Typed filter items",
  sourceRange: "Data!A1:B5",
  targetRange: "A1",
  rowFields: ["Key"],
  valueFields: [{ field: "Sales", name: "Sales" }],
  filters: [{ field: "Key", include: [true, null] }],
  columnGrandTotals: true,
});
assert.deepEqual(typedFilteredPivot.computedValues(), [
  ["Key", "Sales"],
  [true, 20],
  [null, 30],
  ["Grand Total", 50],
]);
const typedFilteredPivotXlsx = await SpreadsheetFile.exportXlsx(typedFilteredPivotWorkbook);
const importedTypedFilteredPivot = (await SpreadsheetFile.importXlsx(typedFilteredPivotXlsx)).worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(importedTypedFilteredPivot.filters, [{ field: "Key", include: [true, null] }]);
assert.deepEqual(importedTypedFilteredPivot.computedValues(), typedFilteredPivot.computedValues());

const emptyFilteredPivotWorkbook = Workbook.create();
const emptyFilteredPivotSheet = emptyFilteredPivotWorkbook.worksheets.add("Data");
emptyFilteredPivotSheet.getRange("A1:B3").values = [["Region", "Sales"], ["East", 10], ["West", 20]];
emptyFilteredPivotSheet.pivotTables.add({
  sourceRange: "A1:B3",
  targetRange: "D1",
  rowFields: ["Region"],
  valueFields: [{ field: "Sales" }],
  filters: [{ field: "Region", exclude: ["East", "West"] }],
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(emptyFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_filter" && /hide every source row/i.test(error.message),
);

const dateFilteredPivotWorkbook = Workbook.create();
const dateFilteredPivotData = dateFilteredPivotWorkbook.worksheets.add("Data");
dateFilteredPivotData.getRange("A1:B5").values = [
  ["Order date", "Sales"],
  [new Date("2026-06-30T00:00:00Z"), 5],
  [new Date("2026-07-01T00:00:00Z"), 10],
  [new Date("2026-07-15T00:00:00Z"), 20],
  [new Date("2026-08-01T00:00:00Z"), 40],
];
const dateFilteredPivotSummary = dateFilteredPivotWorkbook.worksheets.add("Summary");
const dateFilteredPivot = dateFilteredPivotSummary.pivotTables.add({
  name: "July sales",
  sourceRange: "Data!A1:B5",
  targetRange: "A1",
  rowFields: ["Order date"],
  valueFields: [{ field: "Sales" }],
  filters: [{ field: "Order date", type: "dateBetween", value1: "2026-07-01", value2: "2026-07-31" }],
  columnGrandTotals: true,
});
assert.deepEqual(dateFilteredPivot.computedValues(), [
  ["Order date", "sum of Sales"],
  [new Date("2026-07-01T00:00:00Z"), 10],
  [new Date("2026-07-15T00:00:00Z"), 20],
  ["Grand Total", 30],
]);
const dateFilteredPivotXlsx = await SpreadsheetFile.exportXlsx(dateFilteredPivotWorkbook);
const dateFilteredPivotZip = await JSZip.loadAsync(new Uint8Array(await dateFilteredPivotXlsx.arrayBuffer()));
const dateFilteredPivotPart = Object.keys(dateFilteredPivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const dateFilteredPivotCache = Object.keys(dateFilteredPivotZip.files).find((name) => /pivotCache\/pivotCacheDefinition.*\.xml$/i.test(name));
const dateFilteredPivotXml = await dateFilteredPivotZip.file(dateFilteredPivotPart).async("text");
const dateFilteredPivotCacheXml = await dateFilteredPivotZip.file(dateFilteredPivotCache).async("text");
assert.match(dateFilteredPivotXml, /<filters count="1"><filter fld="0" type="dateBetween" id="1" stringValue1="2026-07-01T00:00:00" stringValue2="2026-07-31T00:00:00"><autoFilter \/><\/filter><\/filters>/);
assert.match(dateFilteredPivotCacheXml, /containsDate="1"[^>]*minDate="2026-06-30T00:00:00"[^>]*maxDate="2026-08-01T00:00:00"[\s\S]*<d v="2026-07-01T00:00:00"/);
const importedDateFilteredPivotWorkbook = await SpreadsheetFile.importXlsx(dateFilteredPivotXlsx);
const importedDateFilteredPivot = importedDateFilteredPivotWorkbook.worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(importedDateFilteredPivot.filters, [{
  field: "Order date",
  type: "dateBetween",
  value1: "2026-07-01",
  value2: "2026-07-31",
  useWholeDay: true,
}]);
assert.deepEqual(importedDateFilteredPivot.computedValues(), [
  ["Order date", "Sum of Sales"],
  [46_204, 10],
  [46_218, 20],
  ["Grand Total", 30],
]);
const secondDateFilteredPivotXlsx = await SpreadsheetFile.exportXlsx(importedDateFilteredPivotWorkbook);
const secondDateFilteredPivotZip = await JSZip.loadAsync(new Uint8Array(await secondDateFilteredPivotXlsx.arrayBuffer()));
assert.equal(await secondDateFilteredPivotZip.file(dateFilteredPivotPart).async("text"), dateFilteredPivotXml);
importedDateFilteredPivot.filters[0].value2 = "2026-08-31";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(importedDateFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_edit" && /read-only/i.test(error.message),
);

const relativeDateFilteredPivotWorkbook = Workbook.create();
const relativeDateFilteredPivotData = relativeDateFilteredPivotWorkbook.worksheets.add("Data");
relativeDateFilteredPivotData.getRange("A1:B3").values = [["Order date", "Sales"], [new Date("2026-07-19T00:00:00Z"), 10], [new Date("2026-07-20T00:00:00Z"), 20]];
relativeDateFilteredPivotData.pivotTables.add({
  sourceRange: "A1:B3",
  targetRange: "D1",
  rowFields: ["Order date"],
  valueFields: [{ field: "Sales" }],
  filters: [{ field: "Order date", type: "today", asOf: "2026-07-19" }],
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(relativeDateFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_filter" && /relative and sub-day filters remain model-only/i.test(error.message),
);

const date1904PivotWorkbook = Workbook.create({ dateSystem: "1904" });
const date1904PivotSheet = date1904PivotWorkbook.worksheets.add("Data");
date1904PivotSheet.getRange("A1:B2").values = [["Order date", "Sales"], [new Date("2026-07-15T00:00:00Z"), 20]];
date1904PivotSheet.pivotTables.add({
  sourceRange: "A1:B2",
  targetRange: "D1",
  rowFields: ["Order date"],
  valueFields: [{ field: "Sales" }],
  filters: [{ field: "Order date", type: "dateEqual", value: "2026-07-15" }],
});
const date1904PivotXlsx = await SpreadsheetFile.exportXlsx(date1904PivotWorkbook);
const date1904PivotZip = await JSZip.loadAsync(new Uint8Array(await date1904PivotXlsx.arrayBuffer()));
const date1904PivotCache = Object.keys(date1904PivotZip.files).find((name) => /pivotCache\/pivotCacheDefinition.*\.xml$/i.test(name));
assert.match(await date1904PivotZip.file(date1904PivotCache).async("text"), /containsDate="1"[\s\S]*<d v="2026-07-15T00:00:00"/);
assert.deepEqual((await SpreadsheetFile.importXlsx(date1904PivotXlsx)).worksheets.getItem("Data").pivotTables.items[0].filters, [{
  field: "Order date",
  type: "dateEqual",
  value1: "2026-07-15",
  useWholeDay: true,
}]);

for (const filter of [
  { type: "dateNotEqual", value: "2026-07-15" },
  { type: "dateOlderThan", value: "2026-07-15" },
  { type: "dateOlderThanOrEqual", value: "2026-07-15" },
  { type: "dateNewerThan", value: "2026-07-15" },
  { type: "dateNewerThanOrEqual", value: "2026-07-15" },
  { type: "dateNotBetween", value1: "2026-07-01", value2: "2026-07-31" },
]) {
  const conditionWorkbook = Workbook.create();
  const conditionSheet = conditionWorkbook.worksheets.add("Data");
  conditionSheet.getRange("A1:B5").values = [
    ["Order date", "Sales"],
    [new Date("2026-06-30T00:00:00Z"), 5],
    [new Date("2026-07-01T00:00:00Z"), 10],
    [new Date("2026-07-15T00:00:00Z"), 20],
    [new Date("2026-08-01T00:00:00Z"), 40],
  ];
  const conditionPivot = conditionSheet.pivotTables.add({
    name: `${filter.type} sales`,
    sourceRange: "A1:B5",
    targetRange: "D1",
    rowFields: ["Order date"],
    valueFields: [{ field: "Sales" }],
    filters: [{ field: "Order date", ...filter }],
    columnGrandTotals: true,
  });
  const conditionXlsx = await SpreadsheetFile.exportXlsx(conditionWorkbook);
  const conditionZip = await JSZip.loadAsync(new Uint8Array(await conditionXlsx.arrayBuffer()));
  const conditionPivotPart = Object.keys(conditionZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
  assert.match(await conditionZip.file(conditionPivotPart).async("text"), new RegExp(`type="${filter.type}"`));
  const importedConditionPivot = (await SpreadsheetFile.importXlsx(conditionXlsx)).worksheets.getItem("Data").pivotTables.items[0];
  assert.equal(importedConditionPivot.filters[0].type, filter.type);
  assert.equal(importedConditionPivot.filters[0].value1, filter.value1 || filter.value);
  assert.equal(importedConditionPivot.filters[0].value2, filter.value2);
  assert.equal(importedConditionPivot.computedValues().at(-1).at(-1), conditionPivot.computedValues().at(-1).at(-1));
}

const subDayDateFilteredPivotWorkbook = Workbook.create();
const subDayDateFilteredPivotSheet = subDayDateFilteredPivotWorkbook.worksheets.add("Data");
subDayDateFilteredPivotSheet.getRange("A1:B2").values = [["Order date", "Sales"], [new Date("2026-07-15T12:00:00Z"), 20]];
subDayDateFilteredPivotSheet.pivotTables.add({
  sourceRange: "A1:B2",
  targetRange: "D1",
  rowFields: ["Order date"],
  valueFields: [{ field: "Sales" }],
  filters: [{ field: "Order date", type: "dateEqual", value: "2026-07-15T12:00:00Z", useWholeDay: false }],
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(subDayDateFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_filter" && /relative and sub-day filters remain model-only/i.test(error.message),
);

const textDateFilteredPivotWorkbook = Workbook.create();
const textDateFilteredPivotSheet = textDateFilteredPivotWorkbook.worksheets.add("Data");
textDateFilteredPivotSheet.getRange("A1:B2").values = [["Order date", "Sales"], ["2026-07-15", 20]];
textDateFilteredPivotSheet.pivotTables.add({
  sourceRange: "A1:B2",
  targetRange: "D1",
  rowFields: ["Order date"],
  valueFields: [{ field: "Sales" }],
  filters: [{ field: "Order date", type: "dateEqual", value: "2026-07-15" }],
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(textDateFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_filter" && /filter/i.test(error.message),
);

const multiValuePivotWorkbook = Workbook.create();
const multiValuePivotData = multiValuePivotWorkbook.worksheets.add("Data");
multiValuePivotData.getRange("A1:D5").values = [
  ["Region", "Product", "Sales", "Units"],
  ["East", "A", 10, 2],
  ["East", "B", 20, 4],
  ["West", "A", 30, 6],
  ["West", "B", 40, 8],
];
const multiValuePivotSummary = multiValuePivotWorkbook.worksheets.add("Summary");
multiValuePivotSummary.getRange("A1:G4").format = { fill: "#F0FDFA" };
const multiValuePivot = multiValuePivotSummary.pivotTables.add({
  name: "Revenue and units by region",
  sourceRange: "Data!A1:D5",
  targetRange: "A1",
  rowFields: ["Region"],
  columnFields: ["Product"],
  valueFields: [
    { field: "Sales", summarizeBy: "sum", name: "Revenue" },
    { field: "Units", summarizeBy: "average", name: "Average units" },
  ],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
assert.deepEqual(multiValuePivot.computedValues(), [
  ["Region", "A — Revenue", "A — Average units", "B — Revenue", "B — Average units", "Grand Total — Revenue", "Grand Total — Average units"],
  ["East", 10, 2, 20, 4, 30, 3],
  ["West", 30, 6, 40, 8, 70, 7],
  ["Grand Total", 40, 4, 60, 6, 100, 5],
]);
const multiValuePivotXlsx = await SpreadsheetFile.exportXlsx(multiValuePivotWorkbook);
const multiValuePivotZip = await JSZip.loadAsync(new Uint8Array(await multiValuePivotXlsx.arrayBuffer()));
const multiValuePivotPart = Object.keys(multiValuePivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const multiValuePivotCache = Object.keys(multiValuePivotZip.files).find((name) => /pivotCache\/pivotCacheDefinition.*\.xml$/i.test(name));
const multiValuePivotRecords = Object.keys(multiValuePivotZip.files).find((name) => /pivotCache\/pivotCacheRecords.*\.xml$/i.test(name));
const multiValuePivotXml = await multiValuePivotZip.file(multiValuePivotPart).async("text");
assert.match(multiValuePivotXml, /location ref="A1:G4"/);
assert.match(multiValuePivotXml, /colFields count="2">[\s\S]*field x="1"[\s\S]*field x="-2"/);
assert.match(multiValuePivotXml, /dataFields count="2">[\s\S]*name="Revenue"[^>]*fld="2"[^>]*subtotal="sum"[\s\S]*name="Average units"[^>]*fld="3"[^>]*subtotal="average"/);
assert.match(multiValuePivotXml, /colItems count="6">[\s\S]*<i i="1">/);

const importedMultiValuePivotWorkbook = await SpreadsheetFile.importXlsx(multiValuePivotXlsx);
const importedMultiValuePivot = importedMultiValuePivotWorkbook.worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(importedMultiValuePivot.valueFields, [
  { field: "Sales", summarizeBy: "sum", name: "Revenue" },
  { field: "Units", summarizeBy: "average", name: "Average units" },
]);
assert.deepEqual(importedMultiValuePivot.computedValues(), multiValuePivot.computedValues());
assert.equal(importedMultiValuePivotWorkbook.worksheets.getItem("Summary").getRange("G4").format.fill, "#F0FDFA");
const secondMultiValuePivotXlsx = await SpreadsheetFile.exportXlsx(importedMultiValuePivotWorkbook);
const secondMultiValuePivotZip = await JSZip.loadAsync(new Uint8Array(await secondMultiValuePivotXlsx.arrayBuffer()));
assert.equal(await secondMultiValuePivotZip.file(multiValuePivotPart).async("text"), multiValuePivotXml);
assert.equal(await secondMultiValuePivotZip.file(multiValuePivotCache).async("text"), await multiValuePivotZip.file(multiValuePivotCache).async("text"));
assert.equal(await secondMultiValuePivotZip.file(multiValuePivotRecords).async("text"), await multiValuePivotZip.file(multiValuePivotRecords).async("text"));

const hostNormalizedMultiValueZip = await JSZip.loadAsync(new Uint8Array(await multiValuePivotXlsx.arrayBuffer()));
const hostNormalizedMultiValueXml = multiValuePivotXml
  .replace(/<rowItems\b[\s\S]*?<\/rowItems>/, "")
  .replace(/<colItems\b[\s\S]*?<\/colItems>/, "");
assert.notEqual(hostNormalizedMultiValueXml, multiValuePivotXml);
hostNormalizedMultiValueZip.file(multiValuePivotPart, hostNormalizedMultiValueXml);
const hostNormalizedMultiValueBytes = await hostNormalizedMultiValueZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const hostNormalizedMultiValueFile = new FileBlob(hostNormalizedMultiValueBytes, {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  name: "host-normalized-multi-value-pivot.xlsx",
});
const importedHostNormalizedMultiValue = await SpreadsheetFile.importXlsx(hostNormalizedMultiValueFile);
const hostNormalizedPivot = importedHostNormalizedMultiValue.worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(hostNormalizedPivot.valueFields, [
  { field: "Sales", summarizeBy: "sum", name: "Revenue" },
  { field: "Units", summarizeBy: "average", name: "Average units" },
]);
const preservedHostNormalizedMultiValue = await SpreadsheetFile.exportXlsx(importedHostNormalizedMultiValue);
const preservedHostNormalizedMultiValueZip = await JSZip.loadAsync(new Uint8Array(await preservedHostNormalizedMultiValue.arrayBuffer()));
assert.equal(await preservedHostNormalizedMultiValueZip.file(multiValuePivotPart).async("text"), hostNormalizedMultiValueXml);

const noColumnMultiValueWorkbook = Workbook.create();
const noColumnMultiValueData = noColumnMultiValueWorkbook.worksheets.add("Data");
noColumnMultiValueData.getRange("A1:C5").values = [
  ["Region", "Sales", "Units"],
  ["East", 10, 2],
  ["East", 20, 4],
  ["West", 30, 6],
  ["West", 40, 8],
];
const noColumnMultiValueSummary = noColumnMultiValueWorkbook.worksheets.add("Summary");
const noColumnMultiValuePivot = noColumnMultiValueSummary.pivotTables.add({
  name: "Regional metrics",
  sourceRange: "Data!A1:C5",
  targetRange: "A1",
  rowFields: ["Region"],
  valueFields: [
    { field: "Sales", summarizeBy: "sum", name: "Revenue" },
    { field: "Units", summarizeBy: "count", name: "Unit records" },
  ],
  columnGrandTotals: true,
});
assert.deepEqual(noColumnMultiValuePivot.computedValues(), [
  ["Region", "Revenue", "Unit records"],
  ["East", 30, 2],
  ["West", 70, 2],
  ["Grand Total", 100, 4],
]);
const noColumnMultiValueXlsx = await SpreadsheetFile.exportXlsx(noColumnMultiValueWorkbook);
const noColumnMultiValueZip = await JSZip.loadAsync(new Uint8Array(await noColumnMultiValueXlsx.arrayBuffer()));
const noColumnMultiValuePart = Object.keys(noColumnMultiValueZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const noColumnMultiValueXml = await noColumnMultiValueZip.file(noColumnMultiValuePart).async("text");
assert.match(noColumnMultiValueXml, /location ref="A1:C4"/);
assert.match(noColumnMultiValueXml, /colFields count="1">[\s\S]*field x="-2"/);
assert.match(noColumnMultiValueXml, /colItems count="2">[\s\S]*<i i="1">/);
const importedNoColumnMultiValue = await SpreadsheetFile.importXlsx(noColumnMultiValueXlsx);
assert.deepEqual(importedNoColumnMultiValue.worksheets.getItem("Summary").pivotTables.items[0].valueFields, [
  { field: "Sales", summarizeBy: "sum", name: "Revenue" },
  { field: "Units", summarizeBy: "count", name: "Unit records" },
]);

const editedImportedPivot = await SpreadsheetFile.importXlsx(nativePivotXlsx);
editedImportedPivot.worksheets.getItem("Data").getRange("C2").values = [[11]];
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(editedImportedPivot),
  (error) => error?.code === "unsupported_spreadsheet_pivot_edit" && /source data.*read-only/i.test(error.message),
);
const editedPivotOutput = await SpreadsheetFile.importXlsx(nativePivotXlsx);
editedPivotOutput.worksheets.getItem("Summary").getRange("B2").values = [[11]];
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(editedPivotOutput),
  (error) => error?.code === "unsupported_spreadsheet_pivot_edit" && /cached output.*read-only/i.test(error.message),
);
const multiRowPivotWorkbook = Workbook.create();
const multiRowPivotSheet = multiRowPivotWorkbook.worksheets.add("Data");
multiRowPivotSheet.getRange("A1:C3").values = [["Region", "Product", "Sales"], ["East", "A", 10], ["West", "B", 20]];
const multiRowPivot = multiRowPivotSheet.pivotTables.add({
  name: "Sales by region and product",
  sourceRange: "A1:C3",
  targetRange: "E1",
  rowFields: ["Region", "Product"],
  valueFields: [{ field: "Sales", name: "Sales" }],
  columnGrandTotals: true,
});
assert.deepEqual(multiRowPivot.computedValues(), [
  ["Region", "Product", "Sales"],
  ["East", "A", 10],
  ["West", "B", 20],
  ["Grand Total", "", 30],
]);
const multiRowPivotXlsx = await SpreadsheetFile.exportXlsx(multiRowPivotWorkbook);
const multiRowPivotZip = await JSZip.loadAsync(new Uint8Array(await multiRowPivotXlsx.arrayBuffer()));
const multiRowPivotPart = Object.keys(multiRowPivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const multiRowPivotXml = await multiRowPivotZip.file(multiRowPivotPart).async("text");
assert.match(multiRowPivotXml, /location ref="E1:G4"[^>]*firstDataCol="2"/);
assert.match(multiRowPivotXml, /rowFields count="2">[\s\S]*field x="0"[\s\S]*field x="1"/);
assert.match(multiRowPivotXml, /pivotField[^>]*axis="axisRow"[^>]*compact="0"[^>]*defaultSubtotal="0"/);
const importedMultiRowPivotWorkbook = await SpreadsheetFile.importXlsx(multiRowPivotXlsx);
const importedMultiRowPivot = importedMultiRowPivotWorkbook.worksheets.getItem("Data").pivotTables.items[0];
assert.deepEqual(importedMultiRowPivot.rowFields, ["Region", "Product"]);
assert.deepEqual(importedMultiRowPivot.computedValues(), multiRowPivot.computedValues());

const overBudgetRowPivotWorkbook = Workbook.create();
const overBudgetRowPivotSheet = overBudgetRowPivotWorkbook.worksheets.add("Data");
const overBudgetRowHeaders = [...Array.from({ length: 9 }, (_, index) => `Axis ${index + 1}`), "Sales"];
overBudgetRowPivotSheet.getRange("A1:J2").values = [overBudgetRowHeaders, [...Array.from({ length: 9 }, (_, index) => `Value ${index + 1}`), 10]];
overBudgetRowPivotSheet.pivotTables.add({
  sourceRange: "A1:J2",
  targetRange: "L1",
  rowFields: overBudgetRowHeaders.slice(0, 9),
  valueFields: [{ field: "Sales" }],
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(overBudgetRowPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_profile" && /1 through 8 row fields/i.test(error.message),
);
const overBudgetPivotWorkbook = Workbook.create();
const overBudgetPivotSheet = overBudgetPivotWorkbook.worksheets.add("Data");
overBudgetPivotSheet.getRange("A1:B3").values = [["Region", "Sales"], ["East", 10], ["West", 20]];
overBudgetPivotSheet.pivotTables.add({
  sourceRange: "A1:B3",
  targetRange: "D1",
  rowFields: ["Region"],
  valueFields: Array.from({ length: 33 }, (_, index) => ({ field: "Sales", summarizeBy: "sum", name: `Revenue ${index + 1}` })),
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(overBudgetPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_profile" && /1 through 32 value fields/i.test(error.message),
);
const collidingPivotWorkbook = Workbook.create();
const collidingPivotSheet = collidingPivotWorkbook.worksheets.add("Data");
collidingPivotSheet.getRange("A1:C3").values = [["Region", "Product", "Sales"], ["East", "A", 10], ["West", "B", 20]];
collidingPivotSheet.getRange("E1").values = [["occupied"]];
collidingPivotSheet.pivotTables.add({ sourceRange: "A1:C3", targetRange: "E1", rowFields: ["Region"], valueFields: [{ field: "Sales" }] });
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(collidingPivotWorkbook),
  (error) => error?.code === "spreadsheet_pivot_output_collision" && /overlaps existing worksheet cell E1/i.test(error.message),
);
const collidingMultiValuePivotWorkbook = Workbook.create();
const collidingMultiValuePivotData = collidingMultiValuePivotWorkbook.worksheets.add("Data");
collidingMultiValuePivotData.getRange("A1:D3").values = [
  ["Region", "Product", "Sales", "Units"],
  ["East", "A", 10, 2],
  ["West", "B", 20, 4],
];
const collidingMultiValuePivotSummary = collidingMultiValuePivotWorkbook.worksheets.add("Summary");
collidingMultiValuePivotSummary.getRange("G4").values = [["occupied widened edge"]];
collidingMultiValuePivotSummary.pivotTables.add({
  sourceRange: "Data!A1:D3",
  targetRange: "A1",
  rowFields: ["Region"],
  columnFields: ["Product"],
  valueFields: [{ field: "Sales", name: "Revenue" }, { field: "Units", name: "Units" }],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(collidingMultiValuePivotWorkbook),
  (error) => error?.code === "spreadsheet_pivot_output_collision" && /overlaps existing worksheet cell G4/i.test(error.message),
);
const duplicatePivotWorkbook = Workbook.create();
const duplicatePivotData = duplicatePivotWorkbook.worksheets.add("Data");
duplicatePivotData.getRange("A1:B3").values = [["Region", "Sales"], ["East", 10], ["West", 20]];
const duplicatePivotSummaryA = duplicatePivotWorkbook.worksheets.add("Summary A");
const duplicatePivotSummaryB = duplicatePivotWorkbook.worksheets.add("Summary B");
duplicatePivotSummaryA.pivotTables.add({ name: "Sales Pivot", sourceRange: "Data!A1:B3", targetRange: "A1", rowFields: ["Region"], valueFields: [{ field: "Sales" }] });
duplicatePivotSummaryB.pivotTables.add({ name: "sales pivot", sourceRange: "Data!A1:B3", targetRange: "A1", rowFields: ["Region"], valueFields: [{ field: "Sales" }] });
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(duplicatePivotWorkbook),
  (error) => error?.code === "invalid_spreadsheet_pivot" && /name Sales Pivot must be unique across the workbook/i.test(error.message),
);

for (const [summarizeBy, expected] of [
  ["sum", 60],
  ["count", 3],
  ["average", 20],
  ["min", 10],
  ["max", 30],
]) {
  const aggregationWorkbook = Workbook.create();
  const aggregationSheet = aggregationWorkbook.worksheets.add("Data");
  aggregationSheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["East", "A", 20],
    ["East", "A", 30],
  ];
  const aggregationSummary = aggregationWorkbook.worksheets.add("Summary");
  const aggregationPivot = aggregationSummary.pivotTables.add({
    name: `${summarizeBy} Sales`,
    sourceRange: "Data!A1:C4",
    targetRange: "A1",
    rowFields: ["Region"],
    columnFields: ["Product"],
    valueFields: [{ field: "Sales", summarizeBy }],
  });
  assert.equal(aggregationPivot.computedValues()[1][1], expected);
  const aggregationXlsx = await SpreadsheetFile.exportXlsx(aggregationWorkbook);
  const aggregationZip = await JSZip.loadAsync(new Uint8Array(await aggregationXlsx.arrayBuffer()));
  const aggregationPivotPart = Object.keys(aggregationZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
  assert.match(await aggregationZip.file(aggregationPivotPart).async("text"), new RegExp(`subtotal="${summarizeBy}"`));
  const aggregationImported = await SpreadsheetFile.importXlsx(aggregationXlsx);
  const importedAggregationPivot = aggregationImported.worksheets.getItem("Summary").pivotTables.items[0];
  assert.equal(importedAggregationPivot.valueFields[0].summarizeBy, summarizeBy);
  assert.equal(importedAggregationPivot.computedValues()[1][1], expected);
}

const connectionWorkbook = Workbook.create({
  connections: [{ connectionId: 1, name: "Source-free connection", type: 1, refreshedVersion: 1 }],
});
connectionWorkbook.worksheets.add("Main").getRange("A1").values = [["No connection authoring"]];
const connectionPolicyWorkbook = Workbook.create({
  connections: [{ connectionId: 7, name: "Imported-shaped connection", type: 5, refreshedVersion: 8, refreshOnLoad: true }],
});
assert.throws(() => connectionPolicyWorkbook.disableConnectionRefreshOnLoad(8), /does not exist/i);
assert.throws(() => connectionPolicyWorkbook.disableConnectionRefreshOnLoad("connection/0"), /positive integer/i);
assert.deepEqual(connectionPolicyWorkbook.disableConnectionRefreshOnLoad("connection/7"), {
  connectionId: 7,
  name: "Imported-shaped connection",
  type: 5,
  refreshedVersion: 8,
  refreshOnLoad: false,
});
assert.throws(() => connectionPolicyWorkbook.disableConnectionRefreshOnLoad(7), /explicit refreshOnLoad=true/i);
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(connectionWorkbook),
  (error) => error?.code === "unsupported_workbook_features" && /source-free workbook connections/i.test(error.message),
);

const queryWorkbook = Workbook.create();
const querySheet = queryWorkbook.worksheets.add("Main");
querySheet.getRange("A1:B2").values = [["Key", "Value"], ["A", 1]];
const queryTable = querySheet.tables.add("A1:B2", true, "QueryTable");
queryTable.queryTable = { name: "Source-free query", connectionId: 1 };
assert.throws(() => queryTable.setQueryRefreshPolicy({}), /at least one explicit hardening field/i);
assert.throws(() => queryTable.setQueryRefreshPolicy({ disableRefresh: false }), /may only be set to true/i);
assert.throws(() => queryTable.setQueryRefreshPolicy({ backgroundRefresh: true }), /may only be set to false/i);
assert.throws(() => queryTable.setQueryRefreshPolicy({ refresh: {} }), /Unsupported QueryTable refresh policy field/i);
assert.deepEqual(queryTable.setQueryRefreshPolicy({ disableRefresh: true, backgroundRefresh: false }), {
  name: "Source-free query",
  connectionId: 1,
  disableRefresh: true,
  backgroundRefresh: false,
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(queryWorkbook),
  (error) => error?.code === "unsupported_query_table_edit",
);

const dynamicWorkbook = Workbook.create();
const dynamicSheet = dynamicWorkbook.worksheets.add("Main");
const dynamicCell = dynamicSheet.store.get("A1");
dynamicCell.formula = "=SEQUENCE(2)";
dynamicCell.formulaType = "dynamicArray";
dynamicCell.dynamicArrayRef = "A1:A2";
const dynamicXlsx = await SpreadsheetFile.exportXlsx(dynamicWorkbook);
const dynamicZip = await JSZip.loadAsync(new Uint8Array(await dynamicXlsx.arrayBuffer()));
const dynamicMetadataPath = Object.keys(dynamicZip.files).find((name) => /(?:^|\/)metadata(?:[0-9]+)?\.xml$/i.test(name));
assert.ok(dynamicMetadataPath, "source-free dynamic array export must contain a workbook metadata part");
const dynamicMetadataXml = await dynamicZip.file(dynamicMetadataPath).async("text");
assert.match(dynamicMetadataXml, /XLDAPR/);
assert.match(dynamicMetadataXml, /dynamicArrayProperties/);
assert.match(dynamicMetadataXml, /fDynamic="1"/);
const importedDynamicWorkbook = await SpreadsheetFile.importXlsx(dynamicXlsx);
const importedDynamicCell = importedDynamicWorkbook.worksheets.getItem("Main").store.get("A1");
assert.equal(importedDynamicCell.formula, "=SEQUENCE(2)");
assert.equal(importedDynamicCell.formulaType, "dynamicArray");
assert.equal(importedDynamicCell.dynamicArrayRef, "A1:A2");
importedDynamicWorkbook.recalculate();
const recalculatedDynamicCell = importedDynamicWorkbook.worksheets.getItem("Main").store.get("A1");
assert.equal(recalculatedDynamicCell.spillRange, "A1:A2");
assert.deepEqual(importedDynamicWorkbook.worksheets.getItem("Main").getRange("A1:A2").values, [[1], [2]]);
const preservedRecalculatedDynamic = await SpreadsheetFile.exportXlsx(importedDynamicWorkbook);
const preservedRecalculatedImport = await SpreadsheetFile.importXlsx(preservedRecalculatedDynamic);
assert.equal(preservedRecalculatedImport.worksheets.getItem("Main").store.get("A1").formulaType, "dynamicArray");
assert.equal(preservedRecalculatedImport.worksheets.getItem("Main").store.get("A1").dynamicArrayRef, "A1:A2");

const tamperedDynamicWorkbook = await SpreadsheetFile.importXlsx(dynamicXlsx);
tamperedDynamicWorkbook.recalculate();
tamperedDynamicWorkbook.worksheets.getItem("Main").store.get("A2").spillParent = "Main!B1";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(tamperedDynamicWorkbook, { recalculate: false }),
  (error) => error?.code === "unsupported_dynamic_array_edit" && /source-bound and read-only/i.test(error.message),
);
const editedDynamicCacheWorkbook = await SpreadsheetFile.importXlsx(dynamicXlsx);
editedDynamicCacheWorkbook.worksheets.getItem("Main").store.get("A2").value = 99;
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(editedDynamicCacheWorkbook, { recalculate: false }),
  (error) => error?.code === "unsupported_dynamic_array_edit" && /source-bound and read-only/i.test(error.message),
);
const editedDynamicWorkbook = await SpreadsheetFile.importXlsx(dynamicXlsx);
editedDynamicWorkbook.recalculate();
editedDynamicWorkbook.worksheets.getItem("Main").store.get("A1").formula = "=SEQUENCE(3)";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(editedDynamicWorkbook, { recalculate: false }),
  (error) => error?.code === "unsupported_dynamic_array_edit" && /source-bound and read-only/i.test(error.message),
);

const ifsFormulaWorkbook = Workbook.create();
const ifsFormulaSheet = ifsFormulaWorkbook.worksheets.add("Criteria");
ifsFormulaSheet.getRange("A1:C6").values = [
  ["Region", "Amount", "Status"],
  ["East", 14, "Yes"],
  ["East", 7, "Yes"],
  ["West", 3, "Yes"],
  ["East", "n/a", "Yes"],
  ["West", 20, "No"],
];
ifsFormulaSheet.getRange("E1:E4").formulas = [
  ["=MINIFS(B2:B6,A2:A6,\"East\",C2:C6,\"Yes\")"],
  ["=MAXIFS(B2:B6,A2:A6,\"East\",C2:C6,\"Yes\")"],
  ["=MINIFS(B2:B6,A2:A5,\"East\")"],
  ["=MAXIFS(B2:B6,A2:A6,\"North\")"],
];
assert.deepEqual(ifsFormulaSheet.getRange("E1:E4").values, [[7], [14], ["#VALUE!"], [0]]);
ifsFormulaSheet.getRange("G1:G7").formulas = [
  ["=IFS(FALSE,\"wrong\",TRUE,\"selected\")"],
  ["=IFS(FALSE,\"no match\")"],
  ["=IFS(TRUE,\"short circuit\",TRUE,1/0)"],
  ["=SWITCH(A2,\"West\",1,\"East\",2,0)"],
  ["=SWITCH(\"North\",\"West\",1,\"East\",2,0)"],
  ["=SWITCH(\"North\",\"West\",1)"],
  ["=IFS(TRUE)"],
];
assert.deepEqual(ifsFormulaSheet.getRange("G1:G7").values, [["selected"], ["#N/A"], ["short circuit"], [2], [0], ["#N/A"], ["#VALUE!"]]);
const ifsFormulaXlsx = await SpreadsheetFile.exportXlsx(ifsFormulaWorkbook);
const importedIfsFormulaWorkbook = await SpreadsheetFile.importXlsx(ifsFormulaXlsx);
assert.deepEqual(importedIfsFormulaWorkbook.worksheets.getItem("Criteria").getRange("E1:E4").formulas, ifsFormulaSheet.getRange("E1:E4").formulas);
assert.deepEqual(importedIfsFormulaWorkbook.worksheets.getItem("Criteria").getRange("G1:G7").formulas, ifsFormulaSheet.getRange("G1:G7").formulas);

const letFormulaWorkbook = Workbook.create();
const letFormulaSheet = letFormulaWorkbook.worksheets.add("LET bounds");
letFormulaSheet.getRange("A1").values = [[100]];
letFormulaSheet.getRange("C1:C9").formulas = [
  ["=LET(rate,0.1,principal,1000,principal*(1+rate))"],
  ["=LET(base,A1,bonus,5,base+bonus)"],
  ["=LET(x,2,LET(x,3,x)+x)"],
  ["=LET(x,2,IF(x>1,\"selected\",\"wrong\"))"],
  ["=LET(A1,1,A1)"],
  ["=LET(x,SEQUENCE(2),x)"],
  ["=LET(x,1)"],
  ["=LET(x,1,#N/A)"],
  ["=LET(x,A1:A2,x)"],
];
assert.deepEqual(letFormulaSheet.getRange("C1:C9").values, [[1100], [105], [5], ["selected"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#N/A"], ["#VALUE!"]]);
const letFormulaXlsx = await SpreadsheetFile.exportXlsx(letFormulaWorkbook);
const importedLetFormulaWorkbook = await SpreadsheetFile.importXlsx(letFormulaXlsx);
assert.deepEqual(importedLetFormulaWorkbook.worksheets.getItem("LET bounds").getRange("C1:C9").formulas, letFormulaSheet.getRange("C1:C9").formulas);
assert.deepEqual(importedLetFormulaWorkbook.worksheets.getItem("LET bounds").getRange("C1:C9").values, letFormulaSheet.getRange("C1:C9").values);

const averageIfWorkbook = Workbook.create();
const averageIfSheet = averageIfWorkbook.worksheets.add("AVERAGEIF bounds");
averageIfSheet.getRange("A1:B5").values = [
  ["East", 10],
  ["East", "20"],
  ["West", true],
  ["East", null],
  ["East", 30],
];
averageIfSheet.getRange("D1:D4").formulas = [
  ['=AVERAGEIF(A1:A5,"East",B1:B5)'],
  ['=AVERAGEIF(A1:A5,"West",B1:B5)'],
  ['=AVERAGEIFS(B1:B5,A1:A5,"East")'],
  ['=AVERAGEIFS(B1:B5,A1:A4,"East")'],
];
assert.deepEqual(averageIfSheet.getRange("D1:D4").values, [[20], ["#DIV/0!"], [20], ["#VALUE!"]]);
const averageIfXlsx = await SpreadsheetFile.exportXlsx(averageIfWorkbook);
const importedAverageIfWorkbook = await SpreadsheetFile.importXlsx(averageIfXlsx);
assert.deepEqual(importedAverageIfWorkbook.worksheets.getItem("AVERAGEIF bounds").getRange("D1:D4").formulas, averageIfSheet.getRange("D1:D4").formulas);
assert.deepEqual(importedAverageIfWorkbook.worksheets.getItem("AVERAGEIF bounds").getRange("D1:D4").values, [[20], ["#DIV/0!"], [20], ["#VALUE!"]]);

const percentileWorkbook = Workbook.create();
const percentileSheet = percentileWorkbook.worksheets.add("Percentile bounds");
percentileSheet.getRange("A1:A5").values = [[1], [2], [3], [4], [5]];
percentileSheet.getRange("B1:B5").values = [[10], ["20"], [true], [null], [30]];
percentileSheet.getRange("C1:C2").values = [[1], ["#DIV/0!"]];
percentileSheet.getRange("D1:D3").values = [["10"], [true], [null]];
percentileSheet.getRange("E1:E14").formulas = [
  ["=PERCENTILE.INC(A1:A5,0)"],
  ["=PERCENTILE.INC(A1:A5,0.25)"],
  ["=PERCENTILE.INC(A1:A5,0.5)"],
  ["=PERCENTILE.INC(A1:A5,0.75)"],
  ["=PERCENTILE.INC(A1:A5,1)"],
  ["=QUARTILE.INC(A1:A5,1)"],
  ["=QUARTILE.INC(A1:A5,3)"],
  ["=PERCENTILE.INC(B1:B5,0.5)"],
  ["=PERCENTILE.INC(C1:C2,0.5)"],
  ["=PERCENTILE.INC(A1:A5,-0.1)"],
  ["=QUARTILE.INC(A1:A5,1.5)"],
  ["=PERCENTILE.INC(A1:A5)"],
  ["=QUARTILE.INC(A1:A5,5)"],
  ["=PERCENTILE.INC(D1:D3,0.5)"],
];
assert.deepEqual(percentileSheet.getRange("E1:E14").values, [
  [1], [2], [3], [4], [5], [2], [4], [20], ["#DIV/0!"], ["#NUM!"], ["#NUM!"], ["#VALUE!"], ["#NUM!"], ["#NUM!"],
]);
const percentileXlsx = await SpreadsheetFile.exportXlsx(percentileWorkbook);
const importedPercentileWorkbook = await SpreadsheetFile.importXlsx(percentileXlsx);
assert.deepEqual(importedPercentileWorkbook.worksheets.getItem("Percentile bounds").getRange("E1:E14").formulas, percentileSheet.getRange("E1:E14").formulas);
assert.deepEqual(importedPercentileWorkbook.worksheets.getItem("Percentile bounds").getRange("E1:E14").values, percentileSheet.getRange("E1:E14").values);

const mathFormulaWorkbook = Workbook.create();
const mathFormulaSheet = mathFormulaWorkbook.worksheets.add("Math primitives");
mathFormulaSheet.getRange("A1:A3").values = [[2], [3], [-4]];
mathFormulaSheet.getRange("D1:D82").formulas = [
  ["=PRODUCT(A1:A3)"],
  ["=PRODUCT(2,3,4)"],
  ["=PRODUCT()"],
  ["=PRODUCT(2,1/0)"],
  ["=MOD(-3,2)"],
  ["=MOD(3,-2)"],
  ["=MOD(3,0)"],
  ["=POWER(2,8)"],
  ["=POWER(-1,0.5)"],
  ["=SQRT(81)"],
  ["=SQRT(-1)"],
  ["=SIGN(-4)"],
  ["=SIGN(0)"],
  ["=SIGN(4)"],
  ["=PI()"],
  ["=PI(1)"],
  ["=PRODUCT(1E308,1E308)"],
  ["=SUMSQ(A1:A3)"],
  ["=SUMSQ(2,3,4)"],
  ["=SUMSQ()"],
  ["=QUOTIENT(-7,3)"],
  ["=QUOTIENT(7,-3)"],
  ["=QUOTIENT(7,0)"],
  ["=TRUNC(-123.456,2)"],
  ["=TRUNC(123.456,-1)"],
  ["=TRUNC()"],
  ["=RADIANS(180)"],
  ["=DEGREES(PI())"],
  ["=RADIANS()"],
  ["=DEGREES()"],
  ["=SUMSQ(1E308,1E308)"],
  ["=GCD(24,36,48)"],
  ["=GCD(A1:A3)"],
  ["=GCD()"],
  ["=GCD(1E20,2)"],
  ["=LCM(4,6,8)"],
  ["=LCM(0,6)"],
  ["=LCM()"],
  ["=FACT(5)"],
  ["=FACT(0)"],
  ["=FACT(171)"],
  ["=FACT(-1)"],
  ["=FACTDOUBLE(7)"],
  ["=FACTDOUBLE(0)"],
  ["=COMBIN(5,2)"],
  ["=COMBIN(5,6)"],
  ["=COMBINA(3,2)"],
  ["=COMBINA(0,0)"],
  ["=MROUND(10,3)"],
  ["=MROUND(-10,-3)"],
  ["=MROUND(10,-3)"],
  ["=MROUND(10,0)"],
  ["=EVEN(-3.2)"],
  ["=ODD(4.1)"],
  ["=EVEN()"],
  ["=ODD(1E20)"],
  ["=EXP(1)"],
  ["=EXP(1000)"],
  ["=LN(EXP(2))"],
  ["=LN(0)"],
  ["=LOG(1000)"],
  ["=LOG(8,2)"],
  ["=LOG(8,1)"],
  ["=LOG10(1000)"],
  ["=SIN(PI()/2)"],
  ["=COS(PI())"],
  ["=TAN(0)"],
  ["=ASIN(1)"],
  ["=ACOS(-1)"],
  ["=ATAN(1)"],
  ["=ATAN2(1,1)"],
  ["=ATAN2(0,0)"],
  ["=SINH(0.5)"],
  ["=COSH(0)"],
  ["=TANH(0)"],
  ["=ASINH(1)"],
  ["=ACOSH(0)"],
  ["=ATANH(1)"],
  ["=LN()"],
  ["=LOG(8,2,3)"],
  ["=ASIN(2)"],
  ["=ATAN2(1)"],
];
const mathFormulaValues = mathFormulaSheet.getRange("D1:D16").values;
assert.deepEqual(mathFormulaValues.slice(0, 15), [[-24], [24], ["#VALUE!"], ["#DIV/0!"], [1], [-1], ["#DIV/0!"], [256], ["#NUM!"], [9], ["#NUM!"], [-1], [0], [1], [Math.PI]]);
assert.deepEqual(mathFormulaValues[15], ["#VALUE!"]);
assert.deepEqual(mathFormulaSheet.getRange("D17").values, [["#NUM!"]]);
assert.deepEqual(mathFormulaSheet.getRange("D18:D26").values, [[29], [29], ["#VALUE!"], [-2], [-2], ["#DIV/0!"], [-123.45], [120], ["#VALUE!"]]);
assert.deepEqual(mathFormulaSheet.getRange("D27:D30").values, [[Math.PI], [180], ["#VALUE!"], ["#VALUE!"]]);
assert.deepEqual(mathFormulaSheet.getRange("D31").values, [["#NUM!"]]);
assert.deepEqual(mathFormulaSheet.getRange("D32:D56").values, [[12], [1], ["#VALUE!"], ["#NUM!"], [24], [0], ["#VALUE!"], [120], [1], ["#NUM!"], ["#NUM!"], [105], [1], [10], ["#NUM!"], [6], [1], [9], [-9], ["#NUM!"], ["#DIV/0!"], [-4], [5], ["#VALUE!"], ["#NUM!"]]);
assert.equal(mathFormulaSheet.getRange("D57").values[0][0], Math.E);
assert.deepEqual(mathFormulaSheet.getRange("D58:D60").values, [["#NUM!"], [2], ["#NUM!"]]);
assert.ok(Math.abs(mathFormulaSheet.getRange("D61").values[0][0] - 3) < 1e-12);
assert.deepEqual(mathFormulaSheet.getRange("D62:D64").values, [[3], ["#NUM!"], [3]]);
assert.ok(Math.abs(mathFormulaSheet.getRange("D65").values[0][0] - 1) < 1e-12);
assert.ok(Math.abs(mathFormulaSheet.getRange("D66").values[0][0] + 1) < 1e-12);
assert.deepEqual(mathFormulaSheet.getRange("D67").values, [[0]]);
assert.ok(Math.abs(mathFormulaSheet.getRange("D68").values[0][0] - Math.PI / 2) < 1e-12);
assert.ok(Math.abs(mathFormulaSheet.getRange("D69").values[0][0] - Math.PI) < 1e-12);
assert.ok(Math.abs(mathFormulaSheet.getRange("D70:D71").values[0][0] - Math.PI / 4) < 1e-12);
assert.ok(Math.abs(mathFormulaSheet.getRange("D70:D71").values[1][0] - Math.PI / 4) < 1e-12);
assert.deepEqual(mathFormulaSheet.getRange("D72").values, [["#DIV/0!"]]);
assert.ok(Math.abs(mathFormulaSheet.getRange("D73").values[0][0] - Math.sinh(0.5)) < 1e-12);
assert.deepEqual(mathFormulaSheet.getRange("D74:D75").values, [[1], [0]]);
assert.ok(Math.abs(mathFormulaSheet.getRange("D76").values[0][0] - Math.asinh(1)) < 1e-12);
assert.deepEqual(mathFormulaSheet.getRange("D77:D78").values, [["#NUM!"], ["#NUM!"]]);
assert.deepEqual(mathFormulaSheet.getRange("D79:D82").values, [["#VALUE!"], ["#VALUE!"], ["#NUM!"], ["#VALUE!"]]);
const mathFormulaXlsx = await SpreadsheetFile.exportXlsx(mathFormulaWorkbook);
const importedMathFormulaWorkbook = await SpreadsheetFile.importXlsx(mathFormulaXlsx);
const importedMathFormulaSheet = importedMathFormulaWorkbook.worksheets.getItem("Math primitives");
assert.deepEqual(importedMathFormulaSheet.getRange("D1:D82").formulas, mathFormulaSheet.getRange("D1:D82").formulas);
const importedMathFormulaValues = importedMathFormulaSheet.getRange("D1:D82").values;
assert.deepEqual(importedMathFormulaValues.slice(0, 14), mathFormulaValues.slice(0, 14));
assert.ok(Math.abs(importedMathFormulaValues[14][0] - Math.PI) < Number.EPSILON);
assert.deepEqual(importedMathFormulaValues[15], ["#VALUE!"]);
assert.deepEqual(importedMathFormulaValues[16], ["#NUM!"]);
assert.deepEqual(importedMathFormulaValues.slice(17, 26), [[29], [29], ["#VALUE!"], [-2], [-2], ["#DIV/0!"], [-123.45], [120], ["#VALUE!"]]);
assert.ok(Math.abs(importedMathFormulaValues[26][0] - Math.PI) < Number.EPSILON);
assert.deepEqual(importedMathFormulaValues[27], [180]);
assert.deepEqual(importedMathFormulaValues.slice(28, 31), [["#VALUE!"], ["#VALUE!"], ["#NUM!"]]);
assert.deepEqual(importedMathFormulaValues.slice(31, 56), [[12], [1], ["#VALUE!"], ["#NUM!"], [24], [0], ["#VALUE!"], [120], [1], ["#NUM!"], ["#NUM!"], [105], [1], [10], ["#NUM!"], [6], [1], [9], [-9], ["#NUM!"], ["#DIV/0!"], [-4], [5], ["#VALUE!"], ["#NUM!"]]);
assert.equal(importedMathFormulaValues[56][0], Math.E);
assert.deepEqual(importedMathFormulaValues.slice(57, 60), [["#NUM!"], [2], ["#NUM!"]]);
assert.ok(Math.abs(importedMathFormulaValues[60][0] - 3) < 1e-12);
assert.deepEqual(importedMathFormulaValues.slice(61, 64), [[3], ["#NUM!"], [3]]);
assert.ok(Math.abs(importedMathFormulaValues[64][0] - 1) < 1e-12);
assert.ok(Math.abs(importedMathFormulaValues[65][0] + 1) < 1e-12);
assert.deepEqual(importedMathFormulaValues[66], [0]);
assert.ok(Math.abs(importedMathFormulaValues[67][0] - Math.PI / 2) < 1e-12);
assert.ok(Math.abs(importedMathFormulaValues[68][0] - Math.PI) < 1e-12);
assert.ok(Math.abs(importedMathFormulaValues[69][0] - Math.PI / 4) < 1e-12);
assert.ok(Math.abs(importedMathFormulaValues[70][0] - Math.PI / 4) < 1e-12);
assert.deepEqual(importedMathFormulaValues.slice(71, 73), [["#DIV/0!"], [Math.sinh(0.5)]]);
assert.deepEqual(importedMathFormulaValues.slice(73, 75), [[1], [0]]);
assert.ok(Math.abs(importedMathFormulaValues[75][0] - Math.asinh(1)) < 1e-12);
assert.deepEqual(importedMathFormulaValues.slice(76, 82), [["#NUM!"], ["#NUM!"], ["#VALUE!"], ["#VALUE!"], ["#NUM!"], ["#VALUE!"]]);

const formulaIntrospectionWorkbook = Workbook.create();
const formulaIntrospectionSheet = formulaIntrospectionWorkbook.worksheets.add("Formula introspection");
formulaIntrospectionSheet.getRange("A1:C3").values = [[true, "text", 10], [false, null, 20], [null, "", 30]];
formulaIntrospectionSheet.getRange("A3").formulas = [["=1/0"]];
formulaIntrospectionSheet.getRange("E1:E12").formulas = [
  ["=ROWS(A1:C3)"],
  ["=COLUMNS(A1:C3)"],
  ["=ROWS(B2)"],
  ["=COLUMNS(B2)"],
  ["=ISLOGICAL(A1)"],
  ["=ISLOGICAL(B1)"],
  ["=ISNONTEXT(B1)"],
  ["=ISNONTEXT(A2)"],
  ["=ISNONTEXT(A3)"],
  ["=ROWS(A1:A10001)"],
  ["=ROWS(A1)"],
  ["=COLUMNS(A1,B1)"],
];
assert.deepEqual(formulaIntrospectionSheet.getRange("E1:E12").values, [[3], [3], [1], [1], [true], [false], [false], [true], [true], ["#VALUE!"], [1], ["#VALUE!"]]);
const formulaIntrospectionXlsx = await SpreadsheetFile.exportXlsx(formulaIntrospectionWorkbook);
const importedFormulaIntrospectionWorkbook = await SpreadsheetFile.importXlsx(formulaIntrospectionXlsx);
assert.deepEqual(importedFormulaIntrospectionWorkbook.worksheets.getItem("Formula introspection").getRange("E1:E12").formulas, formulaIntrospectionSheet.getRange("E1:E12").formulas);
assert.deepEqual(importedFormulaIntrospectionWorkbook.worksheets.getItem("Formula introspection").getRange("E1:E12").values, formulaIntrospectionSheet.getRange("E1:E12").values);

const dynamicIntrospectionWorkbook = Workbook.create();
const dynamicIntrospectionSheet = dynamicIntrospectionWorkbook.worksheets.add("Dynamic introspection");
dynamicIntrospectionSheet.getRange("A1").formulas = [["=SEQUENCE(2,3)"]];
dynamicIntrospectionSheet.getRange("E1:E2").formulas = [["=ROWS(A1#)"], ["=COLUMNS(A1#)"]];
assert.deepEqual(dynamicIntrospectionSheet.getRange("E1:E2").values, [[2], [3]]);

const addressFormulaWorkbook = Workbook.create();
const addressFormulaSheet = addressFormulaWorkbook.worksheets.add("ADDRESS bounds");
addressFormulaSheet.getRange("A1:A23").formulas = [
  ["=ADDRESS(2,3)"],
  ["=ADDRESS(2,3,2)"],
  ["=ADDRESS(2,3,2,FALSE)"],
  ["=ADDRESS(2,3,1,FALSE,\"[Book1]Sheet1\")"],
  ["=ADDRESS(2,3,1,FALSE,\"EXCEL SHEET\")"],
  ["=ADDRESS(2,3,3,TRUE)"],
  ["=ADDRESS(2,3,4,FALSE)"],
  ["=ADDRESS(1048576,16384)"],
  ["=ADDRESS(2.9,3.9)"],
  ["=ADDRESS(2,3,1,TRUE,\"O'Brien\")"],
  ["=ADDRESS(2,3,1,TRUE,\"A1\")"],
  ["=ADDRESS(2,3,1,TRUE,\"\")"],
  ["=ADDRESS(0,1)"],
  ["=ADDRESS(1,0)"],
  ["=ADDRESS(1048577,1)"],
  ["=ADDRESS(1,16385)"],
  ["=ADDRESS(2,3,5)"],
  ["=ADDRESS(2,3,1,\"bad\")"],
  ["=ADDRESS(#N/A,3)"],
  ["=ADDRESS(2,3,1,TRUE,2)"],
  ["=ADDRESS()"],
  ["=ADDRESS(1)"],
  ["=ADDRESS(1,1,1,TRUE,\"Sheet1\",1)"],
];
assert.deepEqual(addressFormulaSheet.getRange("A1:A23").values, [
  ["$C$2"], ["C$2"], ["R2C[3]"], ["'[Book1]Sheet1'!R2C3"], ["'EXCEL SHEET'!R2C3"],
  ["$C2"], ["R[2]C[3]"], ["$XFD$1048576"], ["$C$2"], ["'O''Brien'!$C$2"], ["A1!$C$2"], ["$C$2"],
  ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#N/A"], ["#VALUE!"],
  ["#VALUE!"], ["#VALUE!"], ["#VALUE!"],
]);
const addressFormulaXlsx = await SpreadsheetFile.exportXlsx(addressFormulaWorkbook);
const importedAddressFormulaWorkbook = await SpreadsheetFile.importXlsx(addressFormulaXlsx);
assert.deepEqual(importedAddressFormulaWorkbook.worksheets.getItem("ADDRESS bounds").getRange("A1:A23").formulas, addressFormulaSheet.getRange("A1:A23").formulas);
assert.deepEqual(importedAddressFormulaWorkbook.worksheets.getItem("ADDRESS bounds").getRange("A1:A23").values, addressFormulaSheet.getRange("A1:A23").values);

const controlFormulaWorkbook = Workbook.create();
const controlFormulaSheet = controlFormulaWorkbook.worksheets.add("Formula controls");
controlFormulaSheet.getRange("A1:A4").values = [[1], [2], [3], [4]];
controlFormulaSheet.getRange("C1:C10").formulas = [
  ["=CHOOSE(1,\"low\",\"medium\",\"high\")"],
  ["=CHOOSE(2.9,\"low\",\"medium\",\"high\")"],
  ["=CHOOSE(A1,\"low\",\"medium\",\"high\")"],
  ["=CHOOSE(0,\"low\",\"medium\")"],
  ["=CHOOSE(4,\"low\",\"medium\")"],
  ["=XOR(TRUE,FALSE,TRUE)"],
  ["=XOR(TRUE,FALSE,FALSE)"],
  ["=XOR(A1>0,A2>0,A3>0,A4>0)"],
  ["=XOR()"],
  ["=CHOOSE(1,1/0,\"selected\")"],
];
assert.deepEqual(controlFormulaSheet.getRange("C1:C10").values, [["low"], ["medium"], ["low"], ["#VALUE!"], ["#VALUE!"], [false], [true], [false], ["#VALUE!"], ["#DIV/0!"]]);
const controlFormulaXlsx = await SpreadsheetFile.exportXlsx(controlFormulaWorkbook);
const importedControlFormulaWorkbook = await SpreadsheetFile.importXlsx(controlFormulaXlsx);
assert.deepEqual(importedControlFormulaWorkbook.worksheets.getItem("Formula controls").getRange("C1:C10").formulas, controlFormulaSheet.getRange("C1:C10").formulas);
assert.deepEqual(importedControlFormulaWorkbook.worksheets.getItem("Formula controls").getRange("C1:C10").values, controlFormulaSheet.getRange("C1:C10").values);

const spillControlWorkbook = Workbook.create();
const spillControlSheet = spillControlWorkbook.worksheets.add("Spill controls");
spillControlSheet.getRange("A1").formulas = [["=SEQUENCE(2)"]];
spillControlSheet.getRange("C1:C2").formulas = [["=CHOOSE(1,A1#,\"selected\")"], ["=XOR(A1#)"]];
spillControlSheet.getRange("C3").formulas = [["=SIN(A1#)"]];
assert.deepEqual(spillControlSheet.getRange("C1:C3").values, [["#VALUE!"], ["#VALUE!"], ["#VALUE!"]]);

const templateFormulaWorkbook = Workbook.create();
const templateFormulaSheet = templateFormulaWorkbook.worksheets.add("Template formulas");
templateFormulaSheet.getRange("A1:A7").values = [[1], [null], [false], ["text"], [0], ["#DIV/0!"], [null]];
templateFormulaSheet.getRange("A7").formulas = [["=IF(FALSE,1,\"\")"]];
templateFormulaSheet.getRange("B1:B5").formulas = [
  ["=COUNTA(A1:A7)"],
  ["=COUNTBLANK(A1:A7)"],
  ["=COUNTBLANK(A1:A2)"],
  ["=COUNTBLANK(A1)"],
  ["=COUNTBLANK(A1:A2,A3)"],
];
assert.deepEqual(templateFormulaSheet.getRange("B1:B5").values, [[6], [2], [1], [0], ["#VALUE!"]]);
templateFormulaSheet.getRange("C1:C6").formulas = [
  ["=TEXT(DATE(2026,7,12),\"yyyymmdd\")"],
  ["=TEXT(DATE(2026,7,12),\"mmm yyyy\")"],
  ["=TEXT(DATE(2026,7,12),\"mmmm yyyy\")"],
  ["=TEXT(DATE(2026,7,2),\"yyyy-mm-dd\")"],
  ["=TEXT(60,\"yyyy-mm-dd\")"],
  ["=TEXT(DATE(2026,7,12),\"0.00\")"],
];
assert.deepEqual(templateFormulaSheet.getRange("C1:C6").values, [["20260712"], ["Jul 2026"], ["July 2026"], ["2026-07-02"], ["1900-02-29"], ["#VALUE!"]]);
templateFormulaSheet.getRange("D1:D2").values = [["20260601|1"], ["20260602|1"]];
templateFormulaSheet.getRange("E1:E4").formulas = [
  ["=MATCH(\"20260531|1\",D1:D2,0)"],
  ["=INDEX(D1:D2,E1)"],
  ["=IFERROR(INDEX(D1:D2,E1),\"not found\")"],
  ["=INDEX(D1:D2,2)"],
];
assert.deepEqual(templateFormulaSheet.getRange("E1:E4").values, [["#N/A"], ["#N/A"], ["not found"], ["20260602|1"]]);
const templateFormula1904Workbook = Workbook.create({ dateSystem: "1904" });
const templateFormula1904Sheet = templateFormula1904Workbook.worksheets.add("Date system");
templateFormula1904Sheet.getRange("A1").formulas = [["=TEXT(DATE(1904,1,1),\"yyyy-mm-dd\")"]];
assert.deepEqual(templateFormula1904Sheet.getRange("A1").values, [["1904-01-01"]]);

const weekNumberWorkbook = Workbook.create();
const weekNumberSheet = weekNumberWorkbook.worksheets.add("Week numbers");
weekNumberSheet.getRange("A1:A16").formulas = [
  ["=WEEKNUM(DATE(2012,3,9))"],
  ["=WEEKNUM(DATE(2012,3,9),2)"],
  ["=WEEKNUM(DATE(2012,3,9),11)"],
  ["=WEEKNUM(DATE(2012,3,9),12)"],
  ["=WEEKNUM(DATE(2012,3,9),17)"],
  ["=WEEKNUM(DATE(2012,3,9),21)"],
  ["=ISOWEEKNUM(DATE(2012,3,9))"],
  ["=ISOWEEKNUM(DATE(2021,1,1))"],
  ["=ISOWEEKNUM(DATE(2021,1,4))"],
  ["=WEEKNUM(DATE(2026,1,1),10)"],
  ["=WEEKNUM(DATE(2026,1,1),2,1)"],
  ["=WEEKNUM()"],
  ["=ISOWEEKNUM()"],
  ["=ISOWEEKNUM(DATE(2026,1,1),2)"],
  ["=ISOWEEKNUM(-1)"],
  ["=ISOWEEKNUM(#N/A)"],
];
assert.deepEqual(weekNumberSheet.getRange("A1:A16").values, [
  [10], [11], [11], [11], [10], [10], [10], [53], [1],
  ["#NUM!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#NUM!"], ["#N/A"],
]);
const weekNumber1904Workbook = Workbook.create({ dateSystem: "1904" });
const weekNumber1904Sheet = weekNumber1904Workbook.worksheets.add("Week numbers");
weekNumber1904Sheet.getRange("A1:A5").formulas = [
  ["=WEEKNUM(DATE(2012,3,9))"],
  ["=WEEKNUM(DATE(2012,3,9),2)"],
  ["=WEEKNUM(DATE(2012,3,9),21)"],
  ["=ISOWEEKNUM(DATE(2021,1,1))"],
  ["=ISOWEEKNUM(DATE(1904,1,1))"],
];
assert.deepEqual(weekNumber1904Sheet.getRange("A1:A5").values, [[10], [11], [10], [53], [53]]);
const weekNumberXlsx = await SpreadsheetFile.exportXlsx(weekNumberWorkbook);
const importedWeekNumberWorkbook = await SpreadsheetFile.importXlsx(weekNumberXlsx);
assert.deepEqual(importedWeekNumberWorkbook.worksheets.getItem("Week numbers").getRange("A1:A16").formulas, weekNumberSheet.getRange("A1:A16").formulas);
assert.deepEqual(importedWeekNumberWorkbook.worksheets.getItem("Week numbers").getRange("A1:A16").values, weekNumberSheet.getRange("A1:A16").values);

const days360Workbook = Workbook.create();
const days360Sheet = days360Workbook.worksheets.add("DAYS360");
days360Sheet.getRange("A1:A17").formulas = [
  ["=DAYS360(DATE(2011,1,30),DATE(2011,2,1))"],
  ["=DAYS360(DATE(2011,1,1),DATE(2011,12,31))"],
  ["=DAYS360(DATE(2011,1,1),DATE(2011,2,1))"],
  ["=DAYS360(DATE(2021,2,28),DATE(2021,3,31))"],
  ["=DAYS360(DATE(2021,2,27),DATE(2021,2,28))"],
  ["=DAYS360(DATE(2021,2,28),DATE(2021,3,31),TRUE)"],
  ["=DAYS360(DATE(2021,1,31),DATE(2021,2,28),TRUE)"],
  ["=DAYS360(DATE(2011,2,1),DATE(2011,1,30))"],
  ["=DAYS360(DATE(2021,1,1),DATE(2021,1,31),FALSE)"],
  ["=DAYS360(DATE(2021,1,1),DATE(2021,1,31),TRUE)"],
  ["=DAYS360(DATE(2021,1,1),DATE(2021,1,31),1)"],
  ["=DAYS360(DATE(2021,1,1),DATE(2021,1,31),\"european\")"],
  ["=DAYS360(-1,1)"],
  ["=DAYS360()"],
  ["=DAYS360(DATE(2021,1,1))"],
  ["=DAYS360(DATE(2021,1,1),DATE(2021,1,31),FALSE,1)"],
  ["=DAYS360(#N/A,DATE(2021,1,31))"],
];
assert.deepEqual(days360Sheet.getRange("A1:A17").values, [
  [1], [360], [30], [30], [4], [32], [28], [-1], [30], [29], [29],
  ["#VALUE!"], ["#NUM!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#N/A"],
]);
const days360Xlsx = await SpreadsheetFile.exportXlsx(days360Workbook);
const importedDays360Workbook = await SpreadsheetFile.importXlsx(days360Xlsx);
assert.deepEqual(importedDays360Workbook.worksheets.getItem("DAYS360").getRange("A1:A17").formulas, days360Sheet.getRange("A1:A17").formulas);
assert.deepEqual(importedDays360Workbook.worksheets.getItem("DAYS360").getRange("A1:A17").values, days360Sheet.getRange("A1:A17").values);
const days3601904Workbook = Workbook.create({ dateSystem: "1904" });
const days3601904Sheet = days3601904Workbook.worksheets.add("DAYS360");
days3601904Sheet.getRange("A1:A4").formulas = [
  ["=DAYS360(DATE(2011,1,30),DATE(2011,2,1))"],
  ["=DAYS360(DATE(2011,1,1),DATE(2011,12,31))"],
  ["=DAYS360(DATE(2021,2,28),DATE(2021,3,31),TRUE)"],
  ["=DAYS360(DATE(1904,1,1),DATE(1904,1,31))"],
];
assert.deepEqual(days3601904Sheet.getRange("A1:A4").values, [[1], [360], [32], [30]]);
const templateFormulaXlsx = await SpreadsheetFile.exportXlsx(templateFormulaWorkbook);
const importedTemplateFormulaWorkbook = await SpreadsheetFile.importXlsx(templateFormulaXlsx);
assert.deepEqual(importedTemplateFormulaWorkbook.worksheets.getItem("Template formulas").getRange("B1:B5").formulas, templateFormulaSheet.getRange("B1:B5").formulas);
assert.deepEqual(importedTemplateFormulaWorkbook.worksheets.getItem("Template formulas").getRange("C1:C6").formulas, templateFormulaSheet.getRange("C1:C6").formulas);

const indexFormulaWorkbook = Workbook.create();
const indexFormulaSheet = indexFormulaWorkbook.worksheets.add("INDEX bounds");
indexFormulaSheet.getRange("A1:C3").values = [[11, 12, 13], [21, 22, 23], [31, 32, 33]];
indexFormulaSheet.getRange("G1:H4").values = [["Key", "Value"], ["Alpha", 101], ["Beta", 202], ["Gamma", 303]];
indexFormulaSheet.tables.add("G1:H4", true, "IndexValues");
indexFormulaWorkbook.definedNames.add("WideIndex", "'INDEX bounds'!A1:A10001");
indexFormulaSheet.getRange("E1:E17").formulas = [
  ["=INDEX(A1:C3,2,3)"],
  ["=INDEX(A1:C3,2)"],
  ["=INDEX(A1:C3,0,2)"],
  ["=INDEX(A1:C3,2,0)"],
  ["=INDEX(A1:C3,1.9,2.9)"],
  ["=INDEX(A1:C3,-1,2)"],
  ["=INDEX(A1:C3,\"x\",2)"],
  ["=INDEX(A1:C3,1,\"x\")"],
  ["=INDEX(A1:C3,#N/A,1)"],
  ["=INDEX(IndexValues[Value],2)"],
  ["=INDEX(A1:C10001,1)"],
  ["=INDEX(WideIndex,1)"],
  ["=INDEX(A1:C3)"],
  ["=INDEX(A1:C3,1,2,3)"],
  ["=INDEX(1,1)"],
  ["=INDEX(A1:C3,4,1)"],
  ["=INDEX(A1:C3,1,4)"],
];
assert.deepEqual(indexFormulaSheet.getRange("E1:E17").values, [
  [23], [21], [12], [21], [12], [12], [12], [11], ["#N/A"], [202],
  ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#REF!"], ["#REF!"], ["#REF!"],
]);
const indexFormulaXlsx = await SpreadsheetFile.exportXlsx(indexFormulaWorkbook);
const importedIndexFormulaWorkbook = await SpreadsheetFile.importXlsx(indexFormulaXlsx);
assert.deepEqual(importedIndexFormulaWorkbook.worksheets.getItem("INDEX bounds").getRange("E1:E17").formulas, indexFormulaSheet.getRange("E1:E17").formulas);

const xlookupFormulaWorkbook = Workbook.create();
const xlookupFormulaSheet = xlookupFormulaWorkbook.worksheets.add("XLOOKUP bounds");
xlookupFormulaSheet.getRange("A1:B6").values = [
  ["Alpha", 101],
  ["Beta", 201],
  ["Beta", 202],
  ["Gamma", 301],
  [10, 401],
  [20, 501],
];
xlookupFormulaSheet.getRange("G1:I2").values = [
  ["East", "West", "North"],
  [11, 22, 33],
];
xlookupFormulaSheet.getRange("D1:D18").formulas = [
  ['=XLOOKUP("Beta",A1:A6,B1:B6)'],
  ['=XLOOKUP("Beta",A1:A6,B1:B6,"missing",0,-1)'],
  ['=XLOOKUP("Gam*",A1:A6,B1:B6,"missing",2)'],
  ['=XLOOKUP(15,A5:A6,B5:B6,"missing",-1)'],
  ['=XLOOKUP(15,A5:A6,B5:B6,"missing",1)'],
  ['=XLOOKUP("missing",A1:A6,B1:B6,"fallback")'],
  ['=XLOOKUP("missing",A1:A6,B1:B6)'],
  ['=XLOOKUP("Beta",A1:A6,B1:B5)'],
  ['=XLOOKUP("Beta",A1:A3,B1:D1)'],
  ['=XLOOKUP("Beta",A1:B3,A1:B3)'],
  ['=XLOOKUP("Beta",A1:A6,B1:B6,"missing",2,2)'],
  ['=XMATCH("Beta",A1:A6,0,-1)'],
  ['=XMATCH("Gam*",A1:A6,2)'],
  ['=XMATCH(15,A5:A6,-1)'],
  ['=XMATCH("Beta",A1:B3)'],
  ['=XMATCH("Beta",A1:A6,0,2)'],
  ['=XLOOKUP("West",G1:I1,G2:I2)'],
  ['=XMATCH("North",G1:I1)'],
];
assert.deepEqual(xlookupFormulaSheet.getRange("D1:D18").values, [
  [201], [202], [301], [401], [501], ["fallback"], ["#N/A"], ["#VALUE!"],
  ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], [3], [4], [1], ["#VALUE!"], ["#VALUE!"], [22], [3],
]);
xlookupFormulaSheet.getRange("E1").formulas = [['=XLOOKUP("Alpha",A1:A10001,B1:B10001)']];
assert.deepEqual(xlookupFormulaSheet.getRange("E1").values, [["#VALUE!"]], "XLOOKUP must not silently scan vectors larger than the formula budget");
const xlookupFormulaXlsx = await SpreadsheetFile.exportXlsx(xlookupFormulaWorkbook);
const importedXlookupFormulaWorkbook = await SpreadsheetFile.importXlsx(xlookupFormulaXlsx);
assert.deepEqual(importedXlookupFormulaWorkbook.worksheets.getItem("XLOOKUP bounds").getRange("D1:D18").formulas, xlookupFormulaSheet.getRange("D1:D18").formulas);

const tableLookupFormulaWorkbook = Workbook.create();
const tableLookupFormulaSheet = tableLookupFormulaWorkbook.worksheets.add("Table lookup bounds");
tableLookupFormulaSheet.getRange("A1:C4").values = [
  [10, "ten", 100],
  [20, "twenty", 200],
  [30, "thirty", 300],
  [40, "forty", 400],
];
tableLookupFormulaSheet.getRange("E1:H3").values = [
  [10, 20, 30, 40],
  [100, 200, 300, 400],
  ["ten", "twenty", "thirty", "forty"],
];
tableLookupFormulaSheet.getRange("J1:L3").values = [
  [20, "twenty", 200],
  [10, "ten", 100],
  [30, "thirty", 300],
];
tableLookupFormulaSheet.getRange("N1:Q3").values = [
  [20, 10, 30, 40],
  [200, 100, 300, 400],
  ["twenty", "ten", "thirty", "forty"],
];
tableLookupFormulaSheet.getRange("S1:T3").values = [["Alpha", 101], ["Beta", 201], ["Gamma", 301]];
tableLookupFormulaSheet.getRange("V1:X2").values = [["East", "West", "North"], [11, 22, 33]];
tableLookupFormulaWorkbook.definedNames.add("WideLookup", "'Table lookup bounds'!A1:C10001");
tableLookupFormulaSheet.getRange("Z1:Z21").formulas = [
  ["=VLOOKUP(20,A1:C4,3,FALSE)"],
  ["=VLOOKUP(25,A1:C4,3)"],
  ["=VLOOKUP(5,A1:C4,3)"],
  ["=VLOOKUP(20,J1:L3,3)"],
  ["=VLOOKUP(20,A1:C4,4,FALSE)"],
  ["=VLOOKUP(20,A1:C4,0,FALSE)"],
  ["=VLOOKUP(20,A1:C4,3,2)"],
  ["=VLOOKUP(\"Bet*\",S1:T3,2,FALSE)"],
  ["=HLOOKUP(20,E1:H3,3,FALSE)"],
  ["=HLOOKUP(25,E1:H3,2)"],
  ["=HLOOKUP(5,E1:H3,2)"],
  ["=HLOOKUP(20,N1:Q3,2)"],
  ["=HLOOKUP(20,E1:H3,4,FALSE)"],
  ["=HLOOKUP(20,E1:H3,0,FALSE)"],
  ["=HLOOKUP(20,E1:H3,2,2)"],
  ["=HLOOKUP(\"West*\",V1:X2,2,FALSE)"],
  ["=VLOOKUP(10,A1:C10001,3,FALSE)"],
  ["=VLOOKUP(10,WideLookup,3,FALSE)"],
  ["=HLOOKUP(20,E1:H3,2,FALSE,99)"],
  ["=VLOOKUP(10,J1:L3,3,FALSE)"],
  ["=HLOOKUP(10,N1:Q3,2,FALSE)"],
];
assert.deepEqual(tableLookupFormulaSheet.getRange("Z1:Z21").values, [
  [200], [200], ["#N/A"], ["#VALUE!"], ["#REF!"], ["#VALUE!"], ["#VALUE!"], [201],
  ["twenty"], [200], ["#N/A"], ["#VALUE!"], ["#REF!"], ["#VALUE!"], ["#VALUE!"], [22], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], [100], [100],
]);
const tableLookupFormulaXlsx = await SpreadsheetFile.exportXlsx(tableLookupFormulaWorkbook);
const importedTableLookupFormulaWorkbook = await SpreadsheetFile.importXlsx(tableLookupFormulaXlsx);
assert.deepEqual(importedTableLookupFormulaWorkbook.worksheets.getItem("Table lookup bounds").getRange("Z1:Z21").formulas, tableLookupFormulaSheet.getRange("Z1:Z21").formulas);

const lookupFormulaWorkbook = Workbook.create();
const lookupFormulaSheet = lookupFormulaWorkbook.worksheets.add("LOOKUP bounds");
lookupFormulaSheet.getRange("A1:B5").values = [
  [4.14, "red"],
  [4.19, "orange"],
  [5.17, "yellow"],
  [5.77, "green"],
  [6.39, "blue"],
];
lookupFormulaSheet.getRange("D1:H2").values = [
  [10, 20, 30, 40, 50],
  ["ten", "twenty", "thirty", "forty", "fifty"],
];
lookupFormulaSheet.getRange("J1:L3").values = [
  [10, "x", "ten"],
  [20, "y", "twenty"],
  [30, "z", "thirty"],
];
lookupFormulaSheet.getRange("N1:Q2").values = [
  [10, 20, 30, 40],
  ["ten", "twenty", "thirty", "forty"],
];
lookupFormulaSheet.getRange("S1:T3").values = [["Alpha", 101], ["Beta", 202], ["Gamma", 303]];
lookupFormulaSheet.getRange("V1:V3").values = [[20], [10], [30]];
lookupFormulaSheet.getRange("W1:W3").values = [[10], ["twenty"], [30]];
lookupFormulaSheet.getRange("X1:X3").values = [[10], ["#N/A"], [30]];
lookupFormulaWorkbook.definedNames.add("WideLookupVector", "'LOOKUP bounds'!A1:A10001");
lookupFormulaSheet.getRange("Z1:Z22").formulas = [
  ["=LOOKUP(4.19,A1:A5,B1:B5)"],
  ["=LOOKUP(5.75,A1:A5,B1:B5)"],
  ["=LOOKUP(7.66,A1:A5,B1:B5)"],
  ["=LOOKUP(0,A1:A5,B1:B5)"],
  ["=LOOKUP(5.75,A1:A5)"],
  ["=LOOKUP(25,D1:H1,D2:H2)"],
  ["=LOOKUP(25,D1:H1,B1:B5)"],
  ["=LOOKUP(25,J1:L3)"],
  ["=LOOKUP(25,N1:Q2)"],
  ["=LOOKUP(5.75,A1:B5)"],
  ["=LOOKUP(\"Beta\",S1:S3,T1:T3)"],
  ["=LOOKUP(\"Betaz\",S1:S3,T1:T3)"],
  ["=LOOKUP(20,V1:V3)"],
  ["=LOOKUP(20,W1:W3)"],
  ["=LOOKUP(30,A1:A5,B1:B4)"],
  ["=LOOKUP(20,A1:B5,B1:B5)"],
  ["=LOOKUP(20,A1:A5,J1:L3)"],
  ["=LOOKUP(20,WideLookupVector)"],
  ["=LOOKUP(#N/A,A1:A5,B1:B5)"],
  ["=LOOKUP(20,X1:X3)"],
  ["=LOOKUP()"],
  ["=LOOKUP(20,A1:A5,B1:B5,1)"],
];
assert.deepEqual(lookupFormulaSheet.getRange("Z1:Z22").values, [
  ["orange"], ["yellow"], ["blue"], ["#N/A"], [5.17], ["twenty"], ["orange"],
  ["twenty"], ["twenty"], ["yellow"], [202], [202], ["#VALUE!"], ["#VALUE!"],
  ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#N/A"], ["#N/A"],
  ["#VALUE!"], ["#VALUE!"],
]);
const lookupFormulaXlsx = await SpreadsheetFile.exportXlsx(lookupFormulaWorkbook);
const importedLookupFormulaWorkbook = await SpreadsheetFile.importXlsx(lookupFormulaXlsx);
assert.deepEqual(importedLookupFormulaWorkbook.worksheets.getItem("LOOKUP bounds").getRange("Z1:Z22").formulas, lookupFormulaSheet.getRange("Z1:Z22").formulas);
assert.deepEqual(importedLookupFormulaWorkbook.worksheets.getItem("LOOKUP bounds").getRange("Z1:Z22").values, lookupFormulaSheet.getRange("Z1:Z22").values);

const matchFormulaWorkbook = Workbook.create();
const matchFormulaSheet = matchFormulaWorkbook.worksheets.add("MATCH bounds");
matchFormulaSheet.getRange("A1:A4").values = [[10], [20], [30], [40]];
matchFormulaSheet.getRange("B1:B4").values = [[40], [30], [20], [10]];
matchFormulaSheet.getRange("C1:C4").values = [[20], [10], [30], [40]];
matchFormulaSheet.getRange("D1:D3").values = [["Alpha"], ["Beta"], ["Gamma"]];
matchFormulaSheet.getRange("E1:E3").values = [[10], ["twenty"], [30]];
matchFormulaSheet.getRange("F1:F3").values = [["Alpha"], ["Beta*"], ["Gamma"]];
matchFormulaSheet.getRange("G1:J1").values = [[10, 20, 30, 40]];
matchFormulaSheet.getRange("L1:M4").values = [["Key", "Value"], ["Alpha", 101], ["Beta", 201], ["Gamma", 301]];
matchFormulaSheet.getRange("N1:N4").values = [[10], [20], [20], [30]];
matchFormulaSheet.getRange("O1:O4").values = [[30], [20], [20], [10]];
matchFormulaSheet.tables.add("L1:M4", true, "MatchKeys");
matchFormulaWorkbook.definedNames.add("WideMatch", "'MATCH bounds'!A1:A10001");
matchFormulaSheet.getRange("Z1:Z21").formulas = [
  ["=MATCH(20,A1:A4,0)"],
  ["=MATCH(25,A1:A4)"],
  ["=MATCH(5,A1:A4,1)"],
  ["=MATCH(25,B1:B4,-1)"],
  ["=MATCH(45,B1:B4,-1)"],
  ["=MATCH(20,C1:C4,1)"],
  ["=MATCH(20,C1:C4,0)"],
  ["=MATCH(\"Bet*\",D1:D3,0)"],
  ["=MATCH(\"Beta~*\",F1:F3,0)"],
  ["=MATCH(\"Beta\",MatchKeys[Key],0)"],
  ["=MATCH(10,A1:B2,0)"],
  ["=MATCH(10,A1:A10001,0)"],
  ["=MATCH(10,WideMatch,0)"],
  ["=MATCH(10,A1:A4,2)"],
  ["=MATCH(10,A1:A4,0,99)"],
  ["=MATCH(20,A1:A4,-1)"],
  ["=MATCH(20,E1:E3,1)"],
  ["=MATCH(30,G1:J1,0)"],
  ["=MATCH(\"missing\",D1:D3,0)"],
  ["=MATCH(20,N1:N4,1)"],
  ["=MATCH(20,O1:O4,-1)"],
];
assert.deepEqual(matchFormulaSheet.getRange("Z1:Z21").values, [
  [2], [2], ["#N/A"], [2], ["#N/A"], ["#VALUE!"], [1], [2], [2], [2],
  ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], [3], ["#N/A"], [3], [3],
]);
const matchFormulaXlsx = await SpreadsheetFile.exportXlsx(matchFormulaWorkbook);
const importedMatchFormulaWorkbook = await SpreadsheetFile.importXlsx(matchFormulaXlsx);
assert.deepEqual(importedMatchFormulaWorkbook.worksheets.getItem("MATCH bounds").getRange("Z1:Z21").formulas, matchFormulaSheet.getRange("Z1:Z21").formulas);

const formulaBudgetWorkbook = Workbook.create();
const formulaBudgetSheet = formulaBudgetWorkbook.worksheets.add("Formula budget");
formulaBudgetSheet.getRange("A1:B1").values = [[1, 2]];
formulaBudgetWorkbook.definedNames.add("WideFormulaBudget", "'Formula budget'!A1:A10001");
formulaBudgetSheet.getRange("F1").values = [["Value"]];
formulaBudgetSheet.tables.add("F1:F10002", true, "FormulaBudgetValues");
formulaBudgetSheet.getRange("D1:D19").formulas = [
  ["=SUM(A1:A10000)"],
  ["=SUM(A1:A10001)"],
  ["=SUM(WideFormulaBudget)"],
  ['=COUNTIF(A1:A10001,">0")'],
  ['=SUMIFS(A1:A10001,A1:A10001,">0")'],
  ["=FILTER(A1:A10001,A1:A10001>0)"],
  ["=SEQUENCE(10001)"],
  ["=EXPAND(A1,10001,1)"],
  ["=HSTACK(A1:A9999,B1)"],
  ["=SUM(A1:A10000,A1:A10000,A1)"],
  ["=SUM(FormulaBudgetValues[Value])"],
  ['="A1:A10001"'],
  ["=XLOOKUP(1,A1:A10000,B1:B10000)"],
  ["=WRAPROWS(A1,10001)"],
  ["=VSTACK(A1:A10000,B1:B10000)"],
  ["=CHOOSECOLS(A1:A10000,1,1)"],
  ['=IFERROR(SUM(A1:A10001),"fallback")'],
  ["=IFERROR(SUM(A1:A10000,A1:A10000,A1),B1)"],
  ["=AND(TRUE,SUM(A1:A10001))"],
];
assert.deepEqual(formulaBudgetSheet.getRange("D1:D19").values, [
  [1], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"],
  ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["A1:A10001"], [2],
  ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["fallback"], [2], ["#VALUE!"],
]);
const formulaBudgetGraph = formulaBudgetWorkbook.formulaGraph({ recalculate: false });
assert.ok(formulaBudgetGraph.errors.some((error) => error.type === "referenceBudgetExceeded" && error.address === "D2" && error.ref === "A1:A10001" && error.requestedCells === 10001 && error.maximumReferenceCells === 10000 && error.maximumFormulaCells === 20000));
assert.ok(formulaBudgetGraph.errors.some((error) => error.type === "referenceBudgetExceeded" && error.address === "D3" && error.ref === "WideFormulaBudget" && error.requestedCells === 10001));
assert.ok(formulaBudgetGraph.errors.some((error) => error.type === "referenceBudgetExceeded" && error.address === "D10" && error.requestedCells === 1 && error.usedCells === 20000 && error.totalCells === 20001));
assert.ok(formulaBudgetGraph.errors.some((error) => error.type === "referenceBudgetExceeded" && error.address === "D11" && error.ref === "FormulaBudgetValues[Value]" && error.requestedCells === 10001));
assert.equal(formulaBudgetGraph.edges.some((edge) => edge.fromAddress === "D2"), false, "oversized ranges must be refused before graph edge expansion");
assert.equal(formulaBudgetGraph.errors.some((error) => error.address === "D12" && error.type === "referenceBudgetExceeded"), false, "cell-like text literals must not consume the formula reference budget");
const formulaBudgetTrace = formulaBudgetWorkbook.trace("'Formula budget'!D2");
assert.deepEqual(formulaBudgetTrace.tree.precedents, [], "trace must refuse an oversized source before walking its cells");
assert.deepEqual(formulaBudgetTrace.tree.referenceBudget, {
  type: "referenceBudgetExceeded",
  ref: "A1:A10001",
  requestedCells: 10001,
  usedCells: 0,
  totalCells: 10001,
  maximumReferenceCells: 10000,
  maximumFormulaCells: 20000,
});
const formulaBudgetInspectRecords = formulaBudgetWorkbook.inspect({ kind: "formula", sheetName: "Formula budget", range: "D2" }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line));
const formulaBudgetInspectRecord = formulaBudgetInspectRecords.find((record) => record.kind === "formula" && record.address === "D2");
assert.deepEqual(formulaBudgetInspectRecord.precedents, [], "inspect must not allocate precedents for a rejected range");
assert.equal(formulaBudgetInspectRecord.referenceBudget?.requestedCells, 10001);
assert.ok(formulaBudgetWorkbook.verify().issues.some((issue) => issue.type === "formulaReferenceBudgetExceeded" && issue.address === "D2"));

const formulaInputBudgetWorkbook = Workbook.create();
const formulaInputBudgetSheet = formulaInputBudgetWorkbook.worksheets.add("Formula input budget");
const tooLongFormula = `=${"1+".repeat(4096)}1`;
const tooDeepFormula = `=${"(".repeat(65)}1${")".repeat(65)}`;
const tooComplexFormula = `=1${"+1".repeat(513)}`;
const tooManyFormulaArguments = `=SUM(1${",1".repeat(513)})`;
const tooManyComparisonOperators = `=1${">=1".repeat(513)}`;
const tooManyPercentOperators = `=1${"%".repeat(513)}`;
const maximumNestingFormula = `=${"(".repeat(64)}1${")".repeat(64)}`;
formulaInputBudgetSheet.getRange("A1:A7").formulas = [[tooLongFormula], [tooDeepFormula], [tooComplexFormula], [tooManyFormulaArguments], [tooManyComparisonOperators], [tooManyPercentOperators], [maximumNestingFormula]];
assert.deepEqual(formulaInputBudgetSheet.getRange("A1:A7").values, [["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"], [1]]);
const formulaInputBudgetGraph = formulaInputBudgetWorkbook.formulaGraph({ recalculate: false });
assert.ok(formulaInputBudgetGraph.errors.some((error) => error.type === "formulaInputBudgetExceeded" && error.address === "A1" && error.reason === "formulaLength" && error.formulaCharacters === 8194 && error.maximumFormulaCharacters === 8192));
assert.ok(formulaInputBudgetGraph.errors.some((error) => error.type === "formulaInputBudgetExceeded" && error.address === "A2" && error.reason === "formulaNesting" && error.nesting === 65 && error.maximumNesting === 64));
assert.ok(formulaInputBudgetGraph.errors.some((error) => error.type === "formulaInputBudgetExceeded" && error.address === "A3" && error.reason === "formulaOperators" && error.operators === 513 && error.maximumOperators === 512));
assert.ok(formulaInputBudgetGraph.errors.some((error) => error.type === "formulaInputBudgetExceeded" && error.address === "A4" && error.reason === "formulaOperators" && error.operators === 513 && error.maximumOperators === 512));
assert.ok(formulaInputBudgetGraph.errors.some((error) => error.type === "formulaInputBudgetExceeded" && error.address === "A5" && error.reason === "formulaOperators" && error.operators === 513 && error.maximumOperators === 512));
assert.ok(formulaInputBudgetGraph.errors.some((error) => error.type === "formulaInputBudgetExceeded" && error.address === "A6" && error.reason === "formulaOperators" && error.operators === 513 && error.maximumOperators === 512));
const formulaInputBudgetTrace = formulaInputBudgetWorkbook.trace("'Formula input budget'!A1");
assert.deepEqual(formulaInputBudgetTrace.tree.precedents, [], "trace must refuse oversized syntax before parsing precedents");
assert.equal(formulaInputBudgetTrace.tree.inputBudget?.reason, "formulaLength");
const formulaInputBudgetInspectRecords = formulaInputBudgetWorkbook.inspect({ kind: "formula", sheetName: "Formula input budget", range: "A2" }).ndjson
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line));
const formulaInputBudgetInspectRecord = formulaInputBudgetInspectRecords.find((record) => record.kind === "formula" && record.address === "A2");
assert.deepEqual(formulaInputBudgetInspectRecord.precedents, [], "inspect must not parse a deep rejected formula");
assert.equal(formulaInputBudgetInspectRecord.inputBudget?.nesting, 65);
assert.ok(formulaInputBudgetWorkbook.verify().issues.some((issue) => issue.type === "formulaInputBudgetExceeded" && issue.address === "A3" && issue.reason === "formulaOperators"));

const expressionFormulaWorkbook = Workbook.create();
const expressionInputs = expressionFormulaWorkbook.worksheets.add("Inputs");
const expressionTextInputs = expressionFormulaWorkbook.worksheets.add("Data & Targets");
const expressionSheet = expressionFormulaWorkbook.worksheets.add("Expressions");
expressionInputs.getRange("A1:A3").values = [[2], [3], ["North"]];
expressionTextInputs.getRange("A1").values = [["North"]];
expressionSheet.getRange("A1:A12").formulas = [
  ["=SUM(Inputs!A1:A2)+1"],
  ["=TEXT(DATE(2026,7,1)+1,\"yyyy-mm-dd\")"],
  ["=IF(MAX(Inputs!A1:A2)^2=9,\"pass\",\"fail\")"],
  ["=\"Source: \"&Inputs!A3"],
  ["=2^3^2"],
  ["=-2^2"],
  ["=1E-7+1"],
  ["=IFERROR(1/0,\"fallback\")"],
  ["=SIN(1)+1"],
  ["=Inputs!A1<=Inputs!A2"],
  ["=Inputs!A1<>Inputs!A2"],
  ["=\"A&B\"&'Data & Targets'!A1"],
];
assert.deepEqual(expressionSheet.getRange("A1:A6").values, [[6], ["2026-07-02"], ["pass"], ["Source: North"], [512], [4]]);
assert.equal(expressionSheet.getRange("A7").values[0][0], 1.0000001);
assert.deepEqual(expressionSheet.getRange("A8:A12").values, [["fallback"], [1 + Math.sin(1)], [true], [true], ["A&BNorth"]]);
expressionSheet.getRange("A13:A21").formulas = [
  ["=50%"],
  ["=50%*2"],
  ["=100%^2"],
  ["=(100^2)%"],
  ["=-50%"],
  ["=50%%"],
  ["=SUM(10%,20%)"],
  ["=SUMPRODUCT(Inputs!A1:A2%*10)"],
  ["=1%/0"],
];
assert.deepEqual(expressionSheet.getRange("A13:A18").values, [[0.5], [1], [1], [100], [-0.5], [0.005]]);
assert.ok(Math.abs(expressionSheet.getRange("A19").values[0][0] - 0.3) < 1e-12);
assert.equal(expressionSheet.getRange("A20").values[0][0], 0.5);
assert.equal(expressionSheet.getRange("A21").values[0][0], "#DIV/0!");
expressionSheet.getRange("B1").values = [[46205]];
expressionSheet.getRange("B1").conditionalFormats.add("expression", {
  formula: "AND(ISNUMBER(B1),B1=DATE(2026,7,1)+1)",
  format: { fill: "#DCFCE7" },
});
const expressionStyles = expressionFormulaWorkbook.inspect({ kind: "computedStyle", sheetName: "Expressions", range: "B1" }).ndjson
  .trim()
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line));
assert.deepEqual(expressionStyles.map((record) => [record.address, record.style.fill]), [["B1", "#DCFCE7"]]);

const typeFormulaWorkbook = Workbook.create();
const typeFormulaSheet = typeFormulaWorkbook.worksheets.add("Formula types");
typeFormulaSheet.getRange("A1:A6").values = [[42], ["hello"], [true], [null], ["#N/A"], ["A1"]];
typeFormulaSheet.getRange("A7").formulas = [["=SEQUENCE(2)"]];
typeFormulaSheet.getRange("B1:B8").formulas = [
  ["=N(A1)"],
  ["=N(A2)"],
  ["=N(A3)"],
  ["=N(A4)"],
  ["=N(A5)"],
  ["=N(1/0)"],
  ["=N(A1:A2)"],
  ["=N(SEQUENCE(2))"],
];
assert.deepEqual(typeFormulaSheet.getRange("B1:B8").values, [[42], [0], [1], [0], ["#N/A"], ["#DIV/0!"], ["#VALUE!"], ["#VALUE!"]]);
typeFormulaSheet.getRange("C1:C7").formulas = [
  ["=T(A1)"],
  ["=T(A2)"],
  ["=T(A3)"],
  ["=T(A4)"],
  ["=T(A5)"],
  ["=T(1/0)"],
  ["=T(A1:A2)"],
];
assert.deepEqual(typeFormulaSheet.getRange("C1:C7").values, [[""], ["hello"], [""], [""], ["#N/A"], ["#DIV/0!"], ["#VALUE!"]]);
typeFormulaSheet.getRange("D1:D10").formulas = [
  ["=TYPE(A1)"],
  ["=TYPE(A2)"],
  ["=TYPE(A3)"],
  ["=TYPE(A4)"],
  ["=TYPE(A5)"],
  ["=TYPE(1/0)"],
  ["=TYPE(A1:A2)"],
  ["=TYPE(A7#)"],
  ["=TYPE(SEQUENCE(2))"],
  ["=TYPE()"],
];
assert.deepEqual(typeFormulaSheet.getRange("D1:D10").values, [[1], [2], [4], [1], [16], [16], [64], [64], [64], ["#VALUE!"]]);
typeFormulaSheet.getRange("E1:E7").formulas = [
  ["=ISREF(A1)"],
  ["=ISREF(A1:A2)"],
  ["=ISREF(A7#)"],
  ["=ISREF(1)"],
  ["=ISREF(\"A1\")"],
  ["=ISREF(SEQUENCE(2))"],
  ["=ISREF()"],
];
assert.deepEqual(typeFormulaSheet.getRange("E1:E7").values, [[true], [true], [true], [false], [false], [false], ["#VALUE!"]]);
const typeFormulaXlsx = await SpreadsheetFile.exportXlsx(typeFormulaWorkbook);
const importedTypeFormulaWorkbook = await SpreadsheetFile.importXlsx(typeFormulaXlsx);
const importedTypeFormulaSheet = importedTypeFormulaWorkbook.worksheets.getItem("Formula types");
assert.deepEqual(importedTypeFormulaSheet.getRange("B1:E7").values, typeFormulaSheet.getRange("B1:E7").values);
assert.deepEqual(importedTypeFormulaSheet.getRange("B1:E7").formulas, typeFormulaSheet.getRange("B1:E7").formulas);

const referenceFormulaWorkbook = Workbook.create();
const referenceFormulaSheet = referenceFormulaWorkbook.worksheets.add("Formula references");
referenceFormulaSheet.getRange("A1").values = [[42]];
referenceFormulaSheet.getRange("A2").formulas = [["=1+1"]];
referenceFormulaSheet.getRange("A3").values = [["plain"]];
referenceFormulaSheet.getRange("A5").formulas = [["=1/0"]];
referenceFormulaSheet.getRange("B1:B7").formulas = [
  ["=ROW()"], ["=COLUMN()"], ["=ROW(A1)"], ["=COLUMN(C5)"],
  ["=ROW(A1:A2)"], ["=COLUMN(SEQUENCE(2))"], ["=ROW()"],
];
assert.deepEqual(referenceFormulaSheet.getRange("B1:B7").values, [[1], [2], [1], [3], ["#VALUE!"], ["#VALUE!"], [7]]);
referenceFormulaSheet.getRange("C1:C5").formulas = [
  ["=ISFORMULA(A1)"], ["=ISFORMULA(A2)"], ["=ISFORMULA(A1:A2)"],
  ["=ISFORMULA(SEQUENCE(2))"], ["=ISFORMULA()"],
];
assert.deepEqual(referenceFormulaSheet.getRange("C1:C5").values, [[false], [true], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"]]);
referenceFormulaSheet.getRange("D1:D6").formulas = [
  ["=FORMULATEXT(A1)"], ["=FORMULATEXT(A2)"], ["=FORMULATEXT(A5)"],
  ["=FORMULATEXT(A1:A2)"], ["=FORMULATEXT(SEQUENCE(2))"], ["=FORMULATEXT()"],
];
assert.deepEqual(referenceFormulaSheet.getRange("D1:D6").values, [["#N/A"], ["=1+1"], ["=1/0"], ["#VALUE!"], ["#VALUE!"], ["#VALUE!"]]);
const referenceFormulaXlsx = await SpreadsheetFile.exportXlsx(referenceFormulaWorkbook);
const importedReferenceFormulaWorkbook = await SpreadsheetFile.importXlsx(referenceFormulaXlsx);
const importedReferenceFormulaSheet = importedReferenceFormulaWorkbook.worksheets.getItem("Formula references");
assert.deepEqual(importedReferenceFormulaSheet.getRange("B1:D6").values, referenceFormulaSheet.getRange("B1:D6").values);
assert.deepEqual(importedReferenceFormulaSheet.getRange("B1:D6").formulas, referenceFormulaSheet.getRange("B1:D6").formulas);

const sumproductMaskWorkbook = Workbook.create();
const sumproductMaskSheet = sumproductMaskWorkbook.worksheets.add("SUMPRODUCT masks");
sumproductMaskSheet.getRange("A1:A4").values = [[100], [200], [300], [400]];
sumproductMaskSheet.getRange("B1:B4").values = [["Open"], [""], ["Closed Won"], ["Open"]];
sumproductMaskSheet.getRange("C1:C4").values = [[10], [20], [30], [40]];
sumproductMaskSheet.getRange("D1:D2").values = [[20], [20]];
sumproductMaskSheet.getRange("E1:E4").formulas = [
  ["=SUMPRODUCT(A1:A4,--(B1:B4<>\"\"),--(C1:C4>=D1),--(C1:C4<=D1+D2),--(B1:B4<>\"Closed Won\"))"],
  ["=SUMPRODUCT(A1:A4,--(B1:B3<>\"\"))"],
  ["=SUMPRODUCT(A1:A4,B1:B4)"],
  ["=SUMPRODUCT(A1:A10001,B1:B10001)"],
];
assert.deepEqual(sumproductMaskSheet.getRange("E1:E4").values, [[400], ["#VALUE!"], [0], ["#VALUE!"]]);

const textPositionWorkbook = Workbook.create();
const textPositionSheet = textPositionWorkbook.worksheets.add("Text position");
textPositionSheet.getRange("A1:A5").values = [
  ["Quarterly Review"],
  ["quarterly review"],
  ["A*B"],
  ["A?B"],
  ["Launch 🚀 Review"],
];
textPositionSheet.getRange("B1:B10").formulas = [
  ["=SEARCH(\"review\",A1)"],
  ["=FIND(\"Review\",A2)"],
  ["=FIND(\"review\",A2)"],
  ["=SEARCH(\"Re*W\",A1)"],
  ["=SEARCH(\"~*\",A3)"],
  ["=SEARCH(\"~?\",A4)"],
  ["=SEARCH(\"🚀\",A5)"],
  ["=FIND(\"*\",A3)"],
  ["=SEARCH(\"Review\",A1,12)"],
  ["=SEARCH(\"R\",A1,99)"],
];
assert.deepEqual(textPositionSheet.getRange("B1:B10").values, [[11], ["#VALUE!"], [11], [11], [2], [2], [8], [2], ["#VALUE!"], ["#VALUE!"]]);
textPositionSheet.getRange("A1:A5").conditionalFormats.add("expression", {
  formula: "NOT(ISERROR(SEARCH(\"review\",A1)))",
  format: { fill: "#FEF3C7" },
});
const textPositionStyles = textPositionWorkbook.inspect({ kind: "computedStyle", sheetName: "Text position", range: "A1:A5" }).ndjson
  .trim()
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line));
assert.deepEqual(textPositionStyles.map((record) => [record.address, record.style.fill]), [["A1", "#FEF3C7"], ["A2", "#FEF3C7"], ["A5", "#FEF3C7"]]);
textPositionSheet.getRange("C1:C3").values = [["Quarterly Review"], ["quarterly review"], ["Draft"]];
textPositionSheet.getRange("C1:C3").conditionalFormats.add("containsText", {
  text: "Review",
  format: { fill: "#DBEAFE" },
});
const containsTextStyles = textPositionWorkbook.inspect({ kind: "computedStyle", sheetName: "Text position", range: "C1:C3" }).ndjson
  .trim()
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line));
assert.deepEqual(containsTextStyles.map((record) => [record.address, record.style.fill]), [["C1", "#DBEAFE"], ["C2", "#DBEAFE"]]);

const textBoundaryWorkbook = Workbook.create();
const textBoundarySheet = textBoundaryWorkbook.worksheets.add("Text boundary");
textBoundarySheet.getRange("A1:A4").values = [
  ["alpha::beta::Gamma"],
  ["ALPHA::beta"],
  ["no delimiter"],
  ["alpha::beta::"],
];
textBoundarySheet.getRange("B1:B14").formulas = [
  ["=TEXTBEFORE(A1,\"::\")"],
  ["=TEXTAFTER(A1,\"::\")"],
  ["=TEXTBEFORE(A1,\"::\",2)"],
  ["=TEXTAFTER(A1,\"::\",-1)"],
  ["=TEXTBEFORE(A1,\"::\",-1)"],
  ["=TEXTAFTER(A2,\"BETA\",1,1)"],
  ["=TEXTBEFORE(A3,\"::\",1,0,1)"],
  ["=TEXTAFTER(A3,\"::\",1,0,1)"],
  ["=TEXTBEFORE(A3,\"::\",1,0,0,\"missing\")"],
  ["=TEXTAFTER(A4,\"::\",-1)"],
  ["=TEXTAFTER(A1,\"\")"],
  ["=TEXTBEFORE(A1,\"::\",0)"],
  ["=TEXTBEFORE(A1:A2,\"::\")"],
  ["=TEXTAFTER(SEQUENCE(2),\"::\")"],
];
assert.deepEqual(textBoundarySheet.getRange("B1:B14").values, [
  ["alpha"],
  ["beta::Gamma"],
  ["alpha::beta"],
  ["Gamma"],
  ["alpha::beta"],
  [""],
  ["no delimiter"],
  [""],
  ["missing"],
  [""],
  ["#VALUE!"],
  ["#VALUE!"],
  ["#VALUE!"],
  ["#VALUE!"],
]);
const textBoundaryXlsx = await SpreadsheetFile.exportXlsx(textBoundaryWorkbook);
const importedTextBoundaryWorkbook = await SpreadsheetFile.importXlsx(textBoundaryXlsx);
assert.deepEqual(importedTextBoundaryWorkbook.worksheets.getItem("Text boundary").getRange("B1:B14").formulas, textBoundarySheet.getRange("B1:B14").formulas);
assert.deepEqual(importedTextBoundaryWorkbook.worksheets.getItem("Text boundary").getRange("B1:B14").values, textBoundarySheet.getRange("B1:B14").values);

const textTransformWorkbook = Workbook.create();
const textTransformSheet = textTransformWorkbook.worksheets.add("Text transforms");
textTransformSheet.getRange("A1:A4").values = [["Alpha"], ["Beta"], ["A\tB\nC\rD\u007F\u0085😀"], ["x".repeat(32_768)]];
textTransformSheet.getRange("B1:B16").formulas = [
  ['=EXACT("Office","office")'],
  ['=EXACT("Office","Office")'],
  ['=EXACT(A1:A2,"x")'],
  ['=REPT("ab",3)'],
  ['=REPT("ab",0)'],
  ['=REPT("ab",-1)'],
  ['=REPT("abc",10923)'],
  ['=REPLACE("abcdef",2,3,"Z")'],
  ['=REPLACE("😀bc",2,1,"X")'],
  ['=REPLACE("abc",99,1,"Z")'],
  ['=REPLACE("abc",0,1,"Z")'],
  ['=SUBSTITUTE("a-b-a","-","/")'],
  ['=SUBSTITUTE("a-b-a","-","/",2)'],
  ['=SUBSTITUTE("a-b","-","/",0)'],
  ['=SUBSTITUTE("a-b","","/")'],
  ['=SUBSTITUTE(A1:A2,"-","/")'],
];
assert.deepEqual(textTransformSheet.getRange("B1:B16").values, [
  [false],
  [true],
  ["#VALUE!"],
  ["ababab"],
  [""],
  ["#VALUE!"],
  ["#VALUE!"],
  ["aZef"],
  ["😀Xc"],
  ["abcZ"],
  ["#VALUE!"],
  ["a/b/a"],
  ["a-b/a"],
  ["#VALUE!"],
  ["#VALUE!"],
  ["#VALUE!"],
]);
textTransformSheet.getRange("B17:B23").formulas = [
  ["=CLEAN(A3)"],
  ["=CLEAN(UNICHAR(9)&\"A\"&UNICHAR(10)&\"B\")"],
  ["=CLEAN(UNICHAR(127)&\"😀\")"],
  ["=CLEAN(\"\")"],
  ["=CLEAN(A1:A2)"],
  ["=CLEAN(A4)"],
  ["=CLEAN(1/0)"],
];
assert.deepEqual(textTransformSheet.getRange("B17:B23").values, [
  ["ABCD\u007F\u0085😀"],
  ["AB"],
  ["\u007F😀"],
  [""],
  ["#VALUE!"],
  ["#VALUE!"],
  ["#DIV/0!"],
]);
const textTransformXlsx = await SpreadsheetFile.exportXlsx(textTransformWorkbook);
const importedTextTransformWorkbook = await SpreadsheetFile.importXlsx(textTransformXlsx);
assert.deepEqual(importedTextTransformWorkbook.worksheets.getItem("Text transforms").getRange("B1:B16").formulas, textTransformSheet.getRange("B1:B16").formulas);
assert.deepEqual(importedTextTransformWorkbook.worksheets.getItem("Text transforms").getRange("B1:B16").values, textTransformSheet.getRange("B1:B16").values);
assert.deepEqual(importedTextTransformWorkbook.worksheets.getItem("Text transforms").getRange("B17:B23").formulas, textTransformSheet.getRange("B17:B23").formulas);
assert.deepEqual(importedTextTransformWorkbook.worksheets.getItem("Text transforms").getRange("B17:B23").values, textTransformSheet.getRange("B17:B23").values);

const textExtractWorkbook = Workbook.create();
const textExtractSheet = textExtractWorkbook.worksheets.add("Text extract");
textExtractSheet.getRange("A1:A4").values = [["😀bc"], ["abc"], ["x".repeat(32_767)], ["x".repeat(32_768)]];
textExtractSheet.getRange("B1:B15").formulas = [
  ["=LEFT(A1)"],
  ["=LEFT(A1,0)"],
  ["=LEFT(A1,2)"],
  ["=RIGHT(A1,2)"],
  ["=RIGHT(A1,0)"],
  ["=MID(A1,2,1)"],
  ["=MID(A1,99,1)"],
  ["=MID(A1,0,1)"],
  ["=MID(A1,2,-1)"],
  ["=LEN(A1)"],
  ["=LEFT(A1:A2,1)"],
  ["=LEFT(A1,1,2)"],
  ["=LEFT(A3,32767)"],
  ["=LEFT(A3,32768)"],
  ["=LEN(A4)"],
];
assert.deepEqual(textExtractSheet.getRange("B1:B12").values, [
  ["😀"],
  [""],
  ["😀b"],
  ["bc"],
  [""],
  ["b"],
  [""],
  ["#VALUE!"],
  ["#VALUE!"],
  [3],
  ["#VALUE!"],
  ["#VALUE!"],
]);
assert.equal(textExtractSheet.getRange("B13").values[0][0].length, 32_767);
assert.deepEqual(textExtractSheet.getRange("B14:B15").values, [["#VALUE!"], ["#VALUE!"]]);
textExtractSheet.getRange("C1:C13").formulas = [
  ["=UNICODE(A1)"],
  ["=UNICODE(\"𐐷\")"],
  ["=UNICODE(\"\")"],
  ["=UNICODE(A1:A2)"],
  ["=UNICODE(\"x\"&\"😀\")"],
  ["=UNICODE(A4)"],
  ["=UNICHAR(128512)"],
  ["=UNICHAR(66615)"],
  ["=UNICHAR(\"65\")"],
  ["=UNICHAR(0)"],
  ["=UNICHAR(55296)"],
  ["=UNICHAR(1114112)"],
  ["=UNICHAR(1/0)"],
];
assert.deepEqual(textExtractSheet.getRange("C1:C13").values, [
  [128512],
  [66615],
  ["#VALUE!"],
  ["#VALUE!"],
  [120],
  ["#VALUE!"],
  ["😀"],
  ["𐐷"],
  ["A"],
  ["#VALUE!"],
  ["#VALUE!"],
  ["#VALUE!"],
  ["#DIV/0!"],
]);
const textExtractXlsx = await SpreadsheetFile.exportXlsx(textExtractWorkbook);
const importedTextExtractWorkbook = await SpreadsheetFile.importXlsx(textExtractXlsx);
const importedTextExtractSheet = importedTextExtractWorkbook.worksheets.getItem("Text extract");
assert.deepEqual(importedTextExtractSheet.getRange("B1:B15").formulas, textExtractSheet.getRange("B1:B15").formulas);
assert.deepEqual(importedTextExtractSheet.getRange("B1:B12").values, textExtractSheet.getRange("B1:B12").values);
assert.equal(importedTextExtractSheet.getRange("B13").values[0][0].length, 32_767);
assert.deepEqual(importedTextExtractSheet.getRange("B14:B15").values, [["#VALUE!"], ["#VALUE!"]]);
assert.deepEqual(importedTextExtractSheet.getRange("C1:C13").formulas, textExtractSheet.getRange("C1:C13").formulas);
assert.deepEqual(importedTextExtractSheet.getRange("C1:C13").values, textExtractSheet.getRange("C1:C13").values);

const textSplitWorkbook = Workbook.create();
const textSplitSheet = textSplitWorkbook.worksheets.add("Text split");
textSplitSheet.getRange("A1:A5").values = [["North|West|South"], ["A||C"], ["a::b\nc"], ["||"], [Array.from({ length: 10_001 }, (_, index) => String(index)).join("|")]];
textSplitSheet.getRange("B1").formulas = [["=TEXTSPLIT(A1,\"|\")"]];
textSplitSheet.getRange("B3").formulas = [["=TEXTSPLIT(A2,\"|\",,TRUE)"]];
textSplitSheet.getRange("B5").formulas = [["=TEXTSPLIT(\"a=1;b=2\",\"=\",\";\")"]];
textSplitSheet.getRange("E1").formulas = [["=TEXTSPLIT(\"aXbxC\",\"x\",,FALSE,1)"]];
textSplitSheet.getRange("E3").formulas = [["=TEXTSPLIT(\"a=1;b=2;c\",\"=\",\";\",FALSE,0,\"missing\")"]];
textSplitSheet.getRange("H1:H5").formulas = [
  ["=TEXTSPLIT(A1,\"\")"],
  ["=TEXTSPLIT(A1:A2,\"|\")"],
  ["=TEXTSPLIT(SEQUENCE(2),\"|\")"],
  ["=TEXTSPLIT(\"||\",\"|\",,TRUE)"],
  ["=TEXTSPLIT(A5,\"|\")"],
];
assert.deepEqual(textSplitSheet.getRange("B1:D1").values, [["North", "West", "South"]]);
assert.deepEqual(textSplitSheet.getRange("B3:C3").values, [["A", "C"]]);
assert.deepEqual(textSplitSheet.getRange("B5:C6").values, [["a", "1"], ["b", "2"]]);
assert.deepEqual(textSplitSheet.getRange("E1:G1").values, [["a", "b", "C"]]);
assert.deepEqual(textSplitSheet.getRange("E3:F5").values, [["a", "1"], ["b", "2"], ["c", "missing"]]);
assert.deepEqual(textSplitSheet.getRange("H1:H5").values, [["#VALUE!"], ["#VALUE!"], ["#VALUE!"], ["#CALC!"], ["#VALUE!"]]);
const textSplitXlsx = await SpreadsheetFile.exportXlsx(textSplitWorkbook);
const importedTextSplitWorkbook = await SpreadsheetFile.importXlsx(textSplitXlsx);
const importedTextSplitSheet = importedTextSplitWorkbook.worksheets.getItem("Text split");
assert.equal(importedTextSplitSheet.getRange("B1").formulas[0][0], "=TEXTSPLIT(A1,\"|\")");
assert.equal(importedTextSplitSheet.getRange("B5").formulas[0][0], "=TEXTSPLIT(\"a=1;b=2\",\"=\",\";\")");
assert.deepEqual(importedTextSplitSheet.getRange("B5:C6").values, [["a", "1"], ["b", "2"]]);

const importedWithoutSourceSnapshot = await SpreadsheetFile.importXlsx(firstXlsx);
const workbookState = importedWithoutSourceSnapshot[Symbol.for("office-kit.workbook-state")];
workbookState.opaqueOpc.sourcePackage = undefined;
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(importedWithoutSourceSnapshot),
  (error) => error?.code === "missing_source_package",
);

console.log("spreadsheet tests passed");
