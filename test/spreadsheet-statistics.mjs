import assert from "node:assert/strict";

import { SpreadsheetFile, Workbook } from "../src/index.mjs";

function assertClose(actual, expected, tolerance = 1e-12) {
  assert.equal(typeof actual, "number");
  assert.ok(Math.abs(actual - expected) <= tolerance, `${actual} should be within ${tolerance} of ${expected}`);
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Statistics");
sheet.getRange("A1:A8").values = [[2], [4], [4], [4], [5], [5], [7], [9]];
sheet.getRange("B1:B5").values = [[10], ["20"], [true], [null], [30]];
sheet.getRange("C1:C2").values = [[1], ["#DIV/0!"]];
sheet.getRange("D1:E6").values = [
  [1, 2],
  [2, 1],
  [3, 4],
  [4, 3],
  [5, 5],
  ["ignored", 99],
];
sheet.getRange("F1:F5").values = [[7], [7], [7], [7], [7]];
sheet.getRange("J1:J3").values = [[1_000_000_000_001], [1_000_000_000_002], [1_000_000_000_003]];
sheet.getRange("K1:K3").values = [[2_000_000_000_002], [2_000_000_000_004], [2_000_000_000_006]];
sheet.getRange("L1:L2").values = [[new Date("2026-08-12T00:00:00.000Z")], [new Date("2026-08-13T00:00:00.000Z")]];
sheet.getRange("L1:L2").format.numberFormat = "yyyy-mm-dd";
sheet.getRange("M1:N8").values = [
  [2, 6],
  [3, 5],
  [9, 11],
  [1, 7],
  [8, 5],
  [7, 4],
  [5, 4],
  ["ignored", 999],
];
sheet.getRange("O1:P5").values = [[6, 20], [7, 28], [9, 31], [15, 38], [21, 40]];
sheet.getRange("H1:H44").formulas = [
  ["=VAR.P(A1:A8)"],
  ["=VAR.S(A1:A8)"],
  ["=STDEV.P(A1:A8)"],
  ["=STDEV.S(A1:A8)"],
  ["=VAR.P(1,2,TRUE,\"4\")"],
  ["=VAR.S(1,2,TRUE,\"4\")"],
  ["=VAR.P(B1:B5)"],
  ["=STDEV.S(B1:B5)"],
  ["=VAR.S(1)"],
  ["=VAR.P()"],
  ["=STDEV.S(\"not a number\",1)"],
  ["=VAR.P(C1:C2)"],
  ["=CORREL(D1:D5,E1:E5)"],
  ["=COVARIANCE.P(D1:D5,E1:E5)"],
  ["=COVARIANCE.S(D1:D5,E1:E5)"],
  ["=CORREL(D1:D6,E1:E6)"],
  ["=CORREL(D1:D6,E1:E5)"],
  ["=CORREL(D1:D5,F1:F5)"],
  ["=COVARIANCE.P(D6,E6)"],
  ["=COVARIANCE.S(D6,E6)"],
  ["=VAR.S(J1:J3)"],
  ["=CORREL(J1:J3,K1:K3)"],
  ["=COVARIANCE.S(J1:J3,K1:K3)"],
  ["=CORREL(D1:D5)"],
  ["=VAR.P(L1:L2)"],
  ["=VAR.P(1,1/0)"],
  ["=SLOPE(M1:M7,N1:N7)"],
  ["=INTERCEPT(M1:M7,N1:N7)"],
  ["=RSQ(M1:M7,N1:N7)"],
  ["=STEYX(M1:M7,N1:N7)"],
  ["=FORECAST.LINEAR(12,M1:M7,N1:N7)"],
  ["=FORECAST.LINEAR(30,O1:O5,P1:P5)"],
  ["=SLOPE(M1:M7,N1:N6)"],
  ["=SLOPE(M1:M2,F1:F2)"],
  ["=RSQ(M1,M1)"],
  ["=STEYX(M1:M2,N1:N2)"],
  ["=FORECAST.LINEAR(\"not a number\",M1:M7,N1:N7)"],
  ["=SLOPE(M1:M8,N1:N8)"],
  ["=RSQ(J1:J3,K1:K3)"],
  ["=SLOPE(J1:J3,K1:K3)"],
  ["=INTERCEPT(J1:J3,K1:K3)"],
  ["=FORECAST.LINEAR(2000000000008,J1:J3,K1:K3)"],
  ["=STEYX(J1:J3,K1:K3)"],
  ["=SLOPE(C1:C2,N1:N2)"],
];

const results = sheet.getRange("H1:H44").values.flat();
assertClose(results[0], 4);
assertClose(results[1], 32 / 7);
assertClose(results[2], 2);
assertClose(results[3], Math.sqrt(32 / 7));
assertClose(results[4], 1.5);
assertClose(results[5], 2);
assertClose(results[6], 100);
assertClose(results[7], Math.sqrt(200));
assert.deepEqual(results.slice(8, 12), ["#DIV/0!", "#VALUE!", "#VALUE!", 0]);
assertClose(results[12], 0.8);
assertClose(results[13], 1.6);
assertClose(results[14], 2);
assertClose(results[15], 0.8);
assert.equal(results[16], "#N/A");
assert.equal(results[17], "#DIV/0!");
assert.equal(results[18], "#DIV/0!");
assert.equal(results[19], "#DIV/0!");
assertClose(results[20], 1);
assertClose(results[21], 1);
assertClose(results[22], 2);
assert.equal(results[23], "#VALUE!");
assertClose(results[24], 0.25);
assert.equal(results[25], "#DIV/0!");
assertClose(results[26], 11 / 36);
assertClose(results[27], 19 / 6);
assertClose(results[28], 121 / 2088);
assertClose(results[29], 3.305718950210041);
assertClose(results[30], 41 / 6);
assertClose(results[31], 10.607253086419755);
assert.deepEqual(results.slice(32, 37), ["#N/A", "#DIV/0!", "#DIV/0!", "#DIV/0!", "#VALUE!"]);
assertClose(results[37], 11 / 36);
assertClose(results[38], 1);
assertClose(results[39], 0.5);
assertClose(results[40], 0);
assertClose(results[41], 1_000_000_000_004);
assertClose(results[42], 0);
assert.equal(results[43], "#DIV/0!");

const spillSheet = workbook.worksheets.add("Spill statistics");
spillSheet.getRange("A1").formulas = [["=SEQUENCE(5)"]];
spillSheet.getRange("C1").formulas = [["=SEQUENCE(5,1,2,2)"]];
spillSheet.getRange("E1:E9").formulas = [
  ["=CORREL(A1#,C1#)"],
  ["=COVARIANCE.P(A1#,C1#)"],
  ["=VAR.P(A1#)"],
  ["=STDEV.S(C1#)"],
  ["=SLOPE(C1#,A1#)"],
  ["=INTERCEPT(C1#,A1#)"],
  ["=RSQ(C1#,A1#)"],
  ["=STEYX(C1#,A1#)"],
  ["=FORECAST.LINEAR(6,C1#,A1#)"],
];
const spillResults = spillSheet.getRange("E1:E9").values.flat();
assertClose(spillResults[0], 1);
assertClose(spillResults[1], 4);
assertClose(spillResults[2], 2);
assertClose(spillResults[3], Math.sqrt(10));
assert.deepEqual(spillResults.slice(4), [2, 0, 1, 0, 12]);

const lineSheet = workbook.worksheets.add("LINEST statistics");
lineSheet.getRange("A1:B7").values = [
  [2, 6],
  [3, 5],
  [9, 11],
  [1, 7],
  [8, 5],
  [7, 4],
  [5, 4],
];
lineSheet.getRange("C1:C3").values = [[1], [1], [1]];
lineSheet.getRange("D1").formulas = [["=LINEST(A1:A7,B1:B7,TRUE,TRUE)"]];
lineSheet.getRange("G1").formulas = [["=LINEST(A1:A7,B1:B7,FALSE,TRUE)"]];
lineSheet.getRange("J1").formulas = [["=LINEST(A1:A7,B1:B7)"]];
lineSheet.getRange("M1").formulas = [["=LINEST(A1:A7,,TRUE,FALSE)"]];
lineSheet.getRange("P1").formulas = [["=LINEST(A1:A3,C1:C3,TRUE,TRUE)"]];
lineSheet.getRange("S1").formulas = [["=LINEST(A1:A7,B1:B7,1,1)"]];
lineSheet.getRange("S7:T7").formulas = [["=TRUE()", "=FALSE()"]];
lineSheet.getRange("V1").formulas = [["=LINEST(A1:A7,B1:B7,TRUE,FALSE)"]];
lineSheet.getRange("Y1:Y5").formulas = [
  ["=LINEST(A1:A7,B1:B6)"],
  ["=LINEST(A1:A7,B1:G1)"],
  ["=LINEST(A1:A7,B1:B7,\"yes\",TRUE)"],
  ["=LINEST(A1:A7,B1:B7,TRUE,TRUE,FALSE)"],
  ["=LINEST(A1:A7,B1:B7,TRUE,TRUE)"],
];
lineSheet.getRange("Y10").formulas = [["=LINEST(A1:B2,A1:B2,TRUE,TRUE)"]];
lineSheet.getRange("Z5").values = [["occupied"]];
workbook.recalculate();

const lineStats = lineSheet.getRange("D1:E5").values;
assertClose(lineStats[0][0], 11 / 36);
assertClose(lineStats[0][1], 19 / 6);
assertClose(lineStats[1][0], 0.5509531583683402);
assertClose(lineStats[1][1], 3.533962208186286);
assertClose(lineStats[2][0], 121 / 2088);
assertClose(lineStats[2][1], 3.305718950210041);
assertClose(lineStats[3][0], 0.30757498729029);
assert.equal(lineStats[3][1], 5);
assertClose(lineStats[4][0], 121 / 36);
assertClose(lineStats[4][1], 1967 / 36);
assert.equal(lineSheet.store.get("D1").spillRange, "D1:E5");
assert.equal(lineSheet.store.get("D1").spillValues.length, 5);

const forcedOrigin = lineSheet.getRange("G1:H5").values;
assertClose(forcedOrigin[0][0], 221 / 288);
assert.equal(forcedOrigin[0][1], 0);
assertClose(forcedOrigin[1][0], 0.191565786320739);
assert.equal(forcedOrigin[1][1], "#N/A");
assertClose(forcedOrigin[2][0], 0.727840367191226);
assertClose(forcedOrigin[2][1], 3.25097919721747);
assertClose(forcedOrigin[3][0], 16.0458851229261);
assert.equal(forcedOrigin[3][1], 6);
assertClose(forcedOrigin[4][0], 169.586805555556);
assertClose(forcedOrigin[4][1], 63.4131944444445);
assertClose(lineSheet.getRange("J1:K1").values[0][0], 11 / 36);
assertClose(lineSheet.getRange("J1:K1").values[0][1], 19 / 6);
assertClose(lineSheet.getRange("V1:W1").values[0][0], 11 / 36);
assertClose(lineSheet.getRange("V1:W1").values[0][1], 19 / 6);
assert.deepEqual(lineSheet.getRange("S1:T5").values, lineStats);
assert.deepEqual(lineSheet.getRange("S7:T7").values, [[true, false]]);

const defaultX = lineSheet.getRange("M1:N1").values[0];
assertClose(defaultX[0], 4 / 7);
assertClose(defaultX[1], 19 / 7);
const removedConstantX = lineSheet.getRange("P1:Q5").values;
assert.equal(removedConstantX[0][0], 0);
assertClose(removedConstantX[0][1], 14 / 3);
assertClose(removedConstantX[1][0], 0);
assertClose(removedConstantX[1][1], Math.sqrt(43 / 9));
assertClose(removedConstantX[2][0], 0);
assertClose(removedConstantX[2][1], Math.sqrt(43 / 3));
assert.deepEqual(removedConstantX[3], ["#N/A", 2]);
assert.equal(removedConstantX[4][0], 0);
assertClose(removedConstantX[4][1], 86 / 3);
assert.deepEqual(lineSheet.getRange("Y1:Y4").values.flat(), ["#N/A", "#N/A", "#VALUE!", "#VALUE!"]);
assert.equal(lineSheet.getRange("Y5").values[0][0], "#SPILL!");
assert.deepEqual(lineSheet.store.get("Y5").spillError, { type: "blocked", addresses: ["Z5"] });
assert.equal(lineSheet.getRange("Y10").values[0][0], "#VALUE!");
lineSheet.getRange("Y5:Z5").clear();

const trendSheet = workbook.worksheets.add("TREND forecast");
trendSheet.getRange("A1:B5").values = [[1, 3], [2, 5], [3, 7], [4, 9], [5, 11]];
trendSheet.getRange("C1:C3").values = [[6], [7], [8]];
trendSheet.getRange("D1").formulas = [["=TREND(B1:B5,A1:A5,C1:C3)"]];
trendSheet.getRange("F1:H1").values = [[6, 7, 8]];
trendSheet.getRange("F2").formulas = [["=TREND(B1:B5,A1:A5,F1:H1)"]];
trendSheet.getRange("J1").formulas = [["=TREND(B1:B5)"]];
trendSheet.getRange("L1").formulas = [["=TREND(B1:B5,A1:A5,,FALSE)"]];
trendSheet.getRange("N1:O3").values = [[1, 2], [1, 4], [1, 6]];
trendSheet.getRange("P1:P2").values = [[2], [3]];
trendSheet.getRange("Q1").formulas = [["=TREND(O1:O3,N1:N3,P1:P2)"]];
trendSheet.getRange("R1:R2").values = [[9], [10]];
trendSheet.getRange("S1").formulas = [["=TREND(D1#,C1:C3,R1:R2)"]];
trendSheet.getRange("U1:U7").formulas = [
  ["=TREND()"],
  ["=TREND(B1:B5,A1:A4,C1:C3)"],
  ["=TREND(A1:B2,A1:B2,C1:C3)"],
  ["=TREND(B1:B5,A1:A5,A1:B2)"],
  ["=TREND(B1:B5,A1:A5,C1:C3,\"yes\")"],
  ["=TREND(B1:B5,A1:A5,V1:V2)"],
  ["=TREND(B1:B5,A1:A5,C1:C3,TRUE,FALSE)"],
];
trendSheet.getRange("V1:V2").values = [[6], ["not numeric"]];
trendSheet.getRange("X1").formulas = [["=TREND(B1:B5,A1:A5,C1:C3)"]];
trendSheet.getRange("X2").values = [["occupied"]];
workbook.recalculate();

assert.deepEqual(trendSheet.getRange("D1:D3").values, [[13], [15], [17]]);
assert.deepEqual(trendSheet.getRange("F2:H2").values, [[13, 15, 17]]);
assert.deepEqual(trendSheet.getRange("J1:J5").values, [[3], [5], [7], [9], [11]]);
const forcedTrend = trendSheet.getRange("L1:L5").values.flat();
for (let index = 0; index < forcedTrend.length; index += 1) assertClose(forcedTrend[index], (25 / 11) * (index + 1));
assert.deepEqual(trendSheet.getRange("Q1:Q2").values, [[4], [4]]);
assert.deepEqual(trendSheet.getRange("S1:S2").values, [[19], [21]]);
assert.deepEqual(trendSheet.getRange("U1:U7").values.flat(), ["#VALUE!", "#N/A", "#VALUE!", "#VALUE!", "#VALUE!", "#VALUE!", "#VALUE!"]);
assert.equal(trendSheet.getRange("X1").values[0][0], "#SPILL!");
assert.deepEqual(trendSheet.store.get("X1").spillError, { type: "blocked", addresses: ["X2"] });
assert.equal(trendSheet.store.get("D1").spillRange, "D1:D3");
assert.equal(trendSheet.store.get("F2").spillRange, "F2:H2");
assert.equal(trendSheet.store.get("J1").spillRange, "J1:J5");
trendSheet.getRange("X1:X3").clear();

const legacyArraySheet = workbook.worksheets.add("Legacy array interop");
legacyArraySheet.getRange("A1:B5").values = [[1, 3], [2, 5], [3, 7], [4, 9], [5, 11]];
legacyArraySheet.getRange("C1:C3").values = [[6], [7], [8]];
legacyArraySheet.getRange("D1").formulas = [["=TREND(B1:B5,A1:A5,C1:C3)"]];
legacyArraySheet.store.get("D1").formulaType = "array";
legacyArraySheet.store.get("D1").arrayRef = "D1:D3";
legacyArraySheet.getRange("D1:D3").values = [[13], [15], [17]];
legacyArraySheet.store.get("D1").formula = "=TREND(B1:B5,A1:A5,C1:C3)";
legacyArraySheet.store.get("D1").formulaType = "array";
legacyArraySheet.store.get("D1").arrayRef = "D1:D3";
legacyArraySheet.getRange("F1").formulas = [["=LINEST(B1:B5,A1:A5,TRUE(),TRUE())"]];
legacyArraySheet.store.get("F1").formulaType = "array";
legacyArraySheet.store.get("F1").arrayRef = "F1:G5";
legacyArraySheet.getRange("F1:G5").values = Array.from({ length: 5 }, () => [0, 0]);
legacyArraySheet.store.get("F1").formula = "=LINEST(B1:B5,A1:A5,TRUE(),TRUE())";
legacyArraySheet.store.get("F1").formulaType = "array";
legacyArraySheet.store.get("F1").arrayRef = "F1:G5";
workbook.recalculate();
assert.deepEqual(legacyArraySheet.getRange("D1:D3").values, [[13], [15], [17]]);
assert.equal(legacyArraySheet.store.get("D1").spillRange, "D1:D3");
assert.deepEqual(legacyArraySheet.getRange("F1:G1").values, [[2, 1]]);
assert.equal(legacyArraySheet.store.get("F1").spillRange, "F1:G5");

const opaqueLegacyWorkbook = Workbook.create();
const opaqueLegacySheet = opaqueLegacyWorkbook.worksheets.add("Opaque");
opaqueLegacySheet.getRange("A1:A3").values = [[1], [2], [3]];
opaqueLegacySheet.store.get("A1").formula = "=SEQUENCE(3)";
opaqueLegacySheet.store.get("A1").formulaType = "array";
opaqueLegacySheet.store.get("A1").arrayRef = "A1:A3";
opaqueLegacyWorkbook.recalculate();
assert.equal(opaqueLegacySheet.getRange("A1").values[0][0], "#SPILL!");
assert.equal(opaqueLegacySheet.getRange("A2").values[0][0], 2);
assert.equal(opaqueLegacySheet.store.get("A2").spillParent, undefined);

const mismatchedLegacyWorkbook = Workbook.create();
const mismatchedLegacySheet = mismatchedLegacyWorkbook.worksheets.add("Mismatched");
mismatchedLegacySheet.getRange("A1:B5").values = [[1, 3], [2, 5], [3, 7], [4, 9], [5, 11]];
mismatchedLegacySheet.getRange("C1:C3").values = [[6], [7], [8]];
mismatchedLegacySheet.getRange("D1:D2").values = [[13], [15]];
Object.assign(mismatchedLegacySheet.store.get("D1"), {
  formula: "=TREND(B1:B5,A1:A5,C1:C3)",
  formulaType: "array",
  arrayRef: "D1:D2",
});
mismatchedLegacyWorkbook.recalculate();
assert.equal(mismatchedLegacySheet.getRange("D1").values[0][0], "#SPILL!");
assert.equal(mismatchedLegacySheet.getRange("D2").values[0][0], 15);
assert.equal(mismatchedLegacySheet.store.get("D2").spillParent, undefined);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
const imported = await SpreadsheetFile.importXlsx(xlsx);
assert.deepEqual(imported.worksheets.getItem("Statistics").getRange("H1:H44").formulas, sheet.getRange("H1:H44").formulas);
assert.deepEqual(imported.worksheets.getItem("Statistics").getRange("H1:H44").values, sheet.getRange("H1:H44").values);
assert.deepEqual(imported.worksheets.getItem("Spill statistics").getRange("E1:E9").formulas, spillSheet.getRange("E1:E9").formulas);
assert.deepEqual(imported.worksheets.getItem("Spill statistics").getRange("E1:E9").values, spillSheet.getRange("E1:E9").values);
const importedLineSheet = imported.worksheets.getItem("LINEST statistics");
assert.deepEqual(importedLineSheet.getRange("D1:E5").values, lineStats);
assert.equal(importedLineSheet.store.get("D1").formulaType, "dynamicArray");
assert.equal(importedLineSheet.store.get("D1").dynamicArrayRef, "D1:E5");
assert.equal(importedLineSheet.store.get("D1").formula, "=LINEST(A1:A7,B1:B7,TRUE,TRUE)");
assert.deepEqual(importedLineSheet.getRange("G1:H5").values, forcedOrigin);
assert.deepEqual(importedLineSheet.getRange("P1:Q5").values, removedConstantX);
const importedTrendSheet = imported.worksheets.getItem("TREND forecast");
assert.equal(importedTrendSheet.getRange("D1").formulas[0][0], "=TREND(B1:B5,A1:A5,C1:C3)");
assert.deepEqual(importedTrendSheet.getRange("D1:D3").values, [[13], [15], [17]]);
assert.equal(importedTrendSheet.store.get("D1").formulaType, "dynamicArray");
assert.equal(importedTrendSheet.store.get("D1").dynamicArrayRef, "D1:D3");
assert.equal(importedTrendSheet.store.get("F2").dynamicArrayRef, "F2:H2");
assert.deepEqual(importedTrendSheet.getRange("S1:S2").values, [[19], [21]]);
const importedLegacyArraySheet = imported.worksheets.getItem("Legacy array interop");
assert.equal(importedLegacyArraySheet.store.get("D1").formulaType, "array");
assert.equal(importedLegacyArraySheet.store.get("D1").arrayRef, "D1:D3");
assert.deepEqual(importedLegacyArraySheet.getRange("D1:D3").values, [[13], [15], [17]]);
assert.equal(importedLegacyArraySheet.store.get("F1").formulaType, "array");
assert.equal(importedLegacyArraySheet.store.get("F1").arrayRef, "F1:G5");
imported.recalculate();
assert.deepEqual(importedLegacyArraySheet.getRange("D1:D3").values, [[13], [15], [17]]);
assert.deepEqual(importedLegacyArraySheet.getRange("F1:G1").values, [[2, 1]]);
importedLineSheet.getRange("A1").values = [[4]];
importedTrendSheet.getRange("B1").values = [[5]];
imported.recalculate();
const updatedLineStats = importedLineSheet.getRange("D1:E5").values;
assert.notEqual(updatedLineStats[0][0], lineStats[0][0]);
const updatedXlsx = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
const updatedRoundTrip = await SpreadsheetFile.importXlsx(updatedXlsx);
assert.deepEqual(updatedRoundTrip.worksheets.getItem("LINEST statistics").getRange("D1:E5").values, updatedLineStats);
assert.equal(updatedRoundTrip.worksheets.getItem("LINEST statistics").store.get("D1").dynamicArrayRef, "D1:E5");
assert.deepEqual(updatedRoundTrip.worksheets.getItem("TREND forecast").getRange("D1:D3").values, importedTrendSheet.getRange("D1:D3").values);
assert.equal(updatedRoundTrip.worksheets.getItem("TREND forecast").store.get("D1").dynamicArrayRef, "D1:D3");

console.log("spreadsheet statistical formula tests passed");
