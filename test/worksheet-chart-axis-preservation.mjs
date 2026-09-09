import assert from "node:assert/strict";
import { Workbook } from "../src/spreadsheet/index.mjs";
import { spreadsheetChartFromWire, spreadsheetChartSnapshot, wireWorksheetCharts } from "../src/codecs/office-kit-spreadsheet-charts.mjs";

const workbook = Workbook.create();
const authoredSheet = workbook.worksheets.add("Authored");
authoredSheet.charts.add("line", {
  title: "Revenue", hasLegend: false,
  categories: ["A", "B"], series: [{ name: "Revenue", values: [2, 4], trendlines: [{ type: "linear" }] }],
  yAxis: { min: 1, max: 100 },
  position: { left: 0, top: 0, width: 400, height: 300 },
});
const source = wireWorksheetCharts(authoredSheet)[0];
source.yAxis.logBase = 10;
source.series[0].trendlines[0].label = { text: "Imported fit", numberFormatCode: "0.00" };
source.source = { editable: true };
const importedSheet = workbook.worksheets.add("Imported");
const chart = spreadsheetChartFromWire(importedSheet, source);
const state = { slots: [{ chart, wire: source, publicSnapshot: spreadsheetChartSnapshot(chart) }] };
assert.equal(wireWorksheetCharts(importedSheet, state)[0].yAxis.logBase, 10);
chart.title = "Edited revenue";
const edited = wireWorksheetCharts(importedSheet, state)[0];
assert.equal(edited.title, "Edited revenue");
assert.equal(edited.yAxis.logBase, 10, "An unrelated chart edit must retain imported logarithmic scaling");
assert.deepEqual(edited.series[0].trendlines[0].label, source.series[0].trendlines[0].label, "An unrelated chart edit must retain imported trendline label state");
console.log("worksheet chart axis and trendline label preservation ok");
