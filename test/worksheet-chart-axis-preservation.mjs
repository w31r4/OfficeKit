import assert from "node:assert/strict";
import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import { SpreadsheetChartTrendlineLabelArtifactSchema, SpreadsheetChartTextStyleArtifactSchema } from "../src/generated/office_kit/artifact/v1/office_artifact_pb.js";
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
source.titleTextStyle = { fontSizePoints: 12, language: "en-US", strike: "noStrike", baselineThousandthPercent: 0, capitalization: "none", letterSpacingHundredthPoints: 0, kerningHundredthPoints: 0, highlightRgb: "000000", shadow: { colorRgb: "112233", rotateWithShape: false, opacityThousandthPercent: 0 }, glow: { colorRgb: "556677", radiusEmu: 0n, opacityThousandthPercent: 0 } };
source.series[0].trendlines[0].label = { numberFormatCode: "0.00", numberFormatLink: 1,
  richText: { paragraphs: [{ runs: [
    { content: { case: "text", value: "Fit " }, style: { bold: true, language: "zh-CN", strike: "dblStrike", baselineThousandthPercent: -25125, capitalization: "small", letterSpacingHundredthPoints: -238, kerningHundredthPoints: 1238, highlightRgb: "FFFF00", shadow: { colorScheme: "accent1" }, glow: { colorScheme: "accent2", radiusEmu: 12700n } } },
    { content: { case: "lineBreak", value: true } },
    { content: { case: "text", value: "A" }, style: { bold: false } },
  ] }] },
  layout: { manual: { xMode: "edge", x: 0, y: -0.125, width: 1.25 } } };
source.source = { editable: true };
const importedSheet = workbook.worksheets.add("Imported");
const chart = spreadsheetChartFromWire(importedSheet, source);
const state = { slots: [{ chart, wire: source, publicSnapshot: spreadsheetChartSnapshot(chart) }] };
assert.equal(wireWorksheetCharts(importedSheet, state)[0].yAxis.logBase, 10);
chart.title = "Edited revenue";
const edited = wireWorksheetCharts(importedSheet, state)[0];
assert.equal(edited.title, "Edited revenue");
assert.equal(edited.titleTextStyle.language, "en-US", "Unrelated edits retain explicit chart language");
assert.equal(edited.titleTextStyle.strike, "noStrike", "Unrelated edits retain explicit strike cancellation");
assert.equal(edited.titleTextStyle.baselineThousandthPercent, 0, "Unrelated edits retain an explicit baseline reset");
assert.equal(edited.yAxis.logBase, 10, "An unrelated chart edit must retain imported logarithmic scaling");
assert.deepEqual(edited.series[0].trendlines[0].label, source.series[0].trendlines[0].label, "An unrelated chart edit must retain imported trendline label state");
for (const layout of [undefined, {}, { manual: {} }, source.series[0].trendlines[0].label.layout]) {
  const message = create(SpreadsheetChartTrendlineLabelArtifactSchema, { layout });
  const roundTrip = fromBinary(SpreadsheetChartTrendlineLabelArtifactSchema, toBinary(SpreadsheetChartTrendlineLabelArtifactSchema, message));
  assert.deepEqual(roundTrip, message, "Layout containers and optional zero must survive protobuf serialization");
}
for (const numberFormatLink of [undefined, 0, 1, 2]) {
  const message = create(SpreadsheetChartTrendlineLabelArtifactSchema, { numberFormatCode: "0", numberFormatLink });
  const roundTrip = fromBinary(SpreadsheetChartTrendlineLabelArtifactSchema, toBinary(SpreadsheetChartTrendlineLabelArtifactSchema, message));
  assert.equal(roundTrip.numberFormatLink, numberFormatLink ?? 0, "Link state must survive wire serialization; legacy default remains zero");
}
const richLabel = create(SpreadsheetChartTrendlineLabelArtifactSchema, source.series[0].trendlines[0].label);
assert.deepEqual(fromBinary(SpreadsheetChartTrendlineLabelArtifactSchema, toBinary(SpreadsheetChartTrendlineLabelArtifactSchema, richLabel)), richLabel,
  "Styled text and ordered breaks must survive wire serialization");
console.log("worksheet chart axis and trendline label preservation ok");
for (const language of [undefined, "en-US", "EN-us", "zh-Hans-CN"]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { language });
  assert.equal(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).language, language);
}
for (const strike of [undefined, "noStrike", "sngStrike", "dblStrike"]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { strike });
  assert.equal(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).strike, strike);
}
for (const baselineThousandthPercent of [undefined, 0, -25125, -400000, 400000]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { baselineThousandthPercent });
  assert.equal(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).baselineThousandthPercent, baselineThousandthPercent);
}

assert.equal(edited.titleTextStyle.capitalization, "none", "Unrelated edits retain explicit capitalization cancellation");
for (const capitalization of [undefined, "none", "small", "all"]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { capitalization });
  assert.equal(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).capitalization, capitalization);
}

assert.equal(edited.titleTextStyle.letterSpacingHundredthPoints, 0, "Unrelated edits retain explicit spacing reset");
for (const letterSpacingHundredthPoints of [undefined, 0, -238, -76800, 76800]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { letterSpacingHundredthPoints });
  assert.equal(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).letterSpacingHundredthPoints, letterSpacingHundredthPoints);
}

assert.equal(edited.titleTextStyle.kerningHundredthPoints, 0, "Unrelated edits retain explicit zero kerning threshold");
for (const kerningHundredthPoints of [undefined, 0, 1238, 76800]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { kerningHundredthPoints });
  assert.equal(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).kerningHundredthPoints, kerningHundredthPoints);
}

assert.equal(edited.titleTextStyle.highlightRgb, "000000", "Unrelated edits retain explicit black highlight");
for (const highlightRgb of [undefined, "000000", "FFFFFF", "FFAACC"]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { highlightRgb });
  assert.equal(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).highlightRgb, highlightRgb);
}

assert.deepEqual(edited.titleTextStyle.shadow, source.titleTextStyle.shadow, "Unrelated chart edits retain shadow and explicit zero/false");
for (const shadow of [undefined, { colorRgb: "112233" }, { colorScheme: "accent1", rotateWithShape: false, blurRadiusEmu: 0n, opacityThousandthPercent: 0 }]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { shadow });
  assert.deepEqual(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).shadow, message.shadow);
}

assert.deepEqual(edited.titleTextStyle.glow, source.titleTextStyle.glow, "Unrelated chart edits retain glow beside shadow and explicit zero radius/opacity");
for (const glow of [undefined, { colorRgb: "112233", radiusEmu: 0n }, { colorScheme: "accent1", radiusEmu: 12700n, opacityThousandthPercent: 0 }]) {
  const message = create(SpreadsheetChartTextStyleArtifactSchema, { glow });
  assert.deepEqual(fromBinary(SpreadsheetChartTextStyleArtifactSchema, toBinary(SpreadsheetChartTextStyleArtifactSchema, message)).glow, message.glow);
}
