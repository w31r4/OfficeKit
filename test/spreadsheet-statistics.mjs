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

const orderSheet = workbook.worksheets.add("Robust statistics");
orderSheet.getRange("A1:A7").values = [[1], [2], [2], [3], [3], [4], [100]];
orderSheet.getRange("B1:B23").formulas = [
  ["=_xlfn.RANK.AVG(2,A1:A7,0)"],
  ["=_xlfn.RANK.AVG(2,A1:A7,1)"],
  ["=_xlfn.RANK.AVG(5,A1:A7,0)"],
  ["=_xlfn.PERCENTILE.EXC(A1:A7,0.5)"],
  ["=_xlfn.PERCENTILE.EXC(A1:A7,0.125)"],
  ["=_xlfn.PERCENTILE.EXC(A1:A7,0.124)"],
  ["=_xlfn.PERCENTILE.EXC(A1:A7,0.875)"],
  ["=_xlfn.PERCENTILE.EXC(A1:A7,0.876)"],
  ["=_xlfn.QUARTILE.EXC(A1:A7,1)"],
  ["=_xlfn.QUARTILE.EXC(A1:A7,2)"],
  ["=_xlfn.QUARTILE.EXC(A1:A7,3)"],
  ["=_xlfn.QUARTILE.EXC(A1:A7,0)"],
  ["=_xlfn.QUARTILE.EXC(A1:A7,4)"],
  ["=_xlfn.QUARTILE.EXC(A1:A7,1.9)"],
  ["=TRIMMEAN(A1:A7,0.2)"],
  ["=TRIMMEAN(A1:A7,0.3)"],
  ["=TRIMMEAN(A1:A7,0.99)"],
  ["=TRIMMEAN(A1:A7,1)"],
  ["=TRIMMEAN(A1:A7,-0.1)"],
  ["=TRIMMEAN(A1:A7,1.1)"],
  ["=MODE.SNGL(1,2,2,3,3)"],
  ["=MODE.SNGL(1,2,3)"],
  ["=MEDIAN(1,TRUE,\"5\")"],
];
orderSheet.getRange("D1").formulas = [["=_xlfn.MODE.MULT(A1:A7)"]];
orderSheet.getRange("F1").formulas = [["=_xlfn.MODE.MULT(1,2,3)"]];
const orderResults = orderSheet.getRange("B1:B23").values.flat();
assert.deepEqual(orderResults.slice(0, 3), [5.5, 2.5, "#N/A"]);
assert.deepEqual(orderResults.slice(3, 8), [3, 1, "#NUM!", 100, "#NUM!"]);
assert.deepEqual(orderResults.slice(8, 14), [2, 3, 4, "#NUM!", "#NUM!", 2]);
assertClose(orderResults[14], 115 / 7);
assertClose(orderResults[15], 2.8);
assert.deepEqual(orderResults.slice(16, 20), [3, 3, "#NUM!", "#NUM!"]);
assert.deepEqual(orderResults.slice(20), [2, "#N/A", 1]);
assert.deepEqual(orderSheet.getRange("D1:D2").values, [[2], [3]]);
assert.equal(orderSheet.store.get("D1").spillRange, "D1:D2");
assert.equal(orderSheet.getRange("F1").values[0][0], "#N/A");

const orderXlsx = await SpreadsheetFile.exportXlsx(workbook);
const importedOrderWorkbook = await SpreadsheetFile.importXlsx(orderXlsx);
const importedOrderSheet = importedOrderWorkbook.worksheets.getItem("Robust statistics");
assert.deepEqual(importedOrderSheet.getRange("B1:B23").values, orderSheet.getRange("B1:B23").values);
assert.deepEqual(importedOrderSheet.getRange("D1:D2").values, [[2], [3]]);
assert.equal(importedOrderSheet.store.get("D1").formula, "=_xlfn.MODE.MULT(A1:A7)");
assert.equal(importedOrderSheet.store.get("D1").formulaType, "dynamicArray");

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

const exponentialSheet = workbook.worksheets.add("Exponential forecast");
exponentialSheet.getRange("A1:B6").values = [[2, 6], [3, 11], [4, 18], [5, 33], [6, 54], [7, 91]];
exponentialSheet.getRange("C1:C3").values = [[8], [9], [10]];
exponentialSheet.getRange("D1").formulas = [["=GROWTH(B1:B6,A1:A6,C1:C3,TRUE)"]];
exponentialSheet.getRange("F1").formulas = [["=GROWTH(B1:B6,A1:A6,C1:C3,FALSE)"]];
exponentialSheet.getRange("H1").formulas = [["=LOGEST(B1:B6,A1:A6,TRUE,TRUE)"]];
exponentialSheet.getRange("K1").formulas = [["=LOGEST(B1:B6,A1:A6,FALSE,FALSE)"]];
exponentialSheet.getRange("M1:Q2").values = [
  [1, 2, 3, 4, 5],
  [6, 12, 24, 48, 96],
];
exponentialSheet.getRange("M3:O3").values = [[6, 7, 8]];
exponentialSheet.getRange("M4").formulas = [["=GROWTH(M2:Q2,M1:Q1,M3:O3)"]];
exponentialSheet.getRange("S1:S3").values = [[2], [8], [32]];
exponentialSheet.getRange("T1:T3").values = [[1], [1], [1]];
exponentialSheet.getRange("U1:U2").values = [[2], [3]];
exponentialSheet.getRange("V1").formulas = [["=GROWTH(S1:S3,T1:T3,U1:U2)"]];
exponentialSheet.getRange("X1:Y1").formulas = [["=GROWTH(B1:B6)", "=LOGEST(B1:B6)"]];
exponentialSheet.getRange("AA1:AA8").formulas = [
  ["=GROWTH()"],
  ["=GROWTH(B1:B6,A1:A5,C1:C3)"],
  ["=GROWTH(A1:B2,A1:B2,C1:C3)"],
  ["=GROWTH(B1:B6,A1:A6,A1:B2)"],
  ["=GROWTH(B1:B6,A1:A6,C1:C3,\"yes\")"],
  ["=GROWTH(B1:B6,A1:A6,AB1:AB2)"],
  ["=LOGEST(AC1:AC3,AD1:AD3)"],
  ["=LOGEST(B1:B6,A1:A6,TRUE,TRUE,FALSE)"],
];
exponentialSheet.getRange("AB1:AB2").values = [[8], ["not numeric"]];
exponentialSheet.getRange("AC1:AD3").values = [[2, 1], [0, 2], [8, 3]];
exponentialSheet.getRange("AF1").formulas = [["=GROWTH(M2:Q2,M1:Q1,AF2)"]];
exponentialSheet.getRange("AF2").values = [[2000]];
exponentialSheet.getRange("AH1").formulas = [["=GROWTH(B1:B6,A1:A6,C1:C3)"]];
exponentialSheet.getRange("AH2").values = [["occupied"]];
workbook.recalculate();

const exponentialGrowth = exponentialSheet.getRange("D1:D3").values.flat();
assertClose(exponentialGrowth[0], 160.27426439333, 1e-10);
assertClose(exponentialGrowth[1], 275.620953867361, 1e-10);
assertClose(exponentialGrowth[2], 473.980713611783, 1e-10);
const forcedGrowth = exponentialSheet.getRange("F1:F3").values.flat();
assertClose(forcedGrowth[0], 241.455433648044, 1e-10);
assertClose(forcedGrowth[1], 479.392909170845, 1e-10);
assertClose(forcedGrowth[2], 951.801158048399, 1e-10);
const logestStats = exponentialSheet.getRange("H1:I5").values;
assertClose(logestStats[0][0], 1.71968316255041, 1e-12);
assertClose(logestStats[0][1], 2.0954450641003, 1e-12);
assertClose(logestStats[1][0], 0.00851365404141612, 1e-14);
assertClose(logestStats[1][1], 0.0409777183642363, 1e-14);
assertClose(logestStats[2][0], 0.999014535947832, 1e-12);
assertClose(logestStats[2][1], 0.0356151700809657, 1e-14);
assertClose(logestStats[3][0], 4055.00143307963, 1e-9);
assert.equal(logestStats[3][1], 4);
assertClose(logestStats[4][0], 5.14352739605477, 1e-12);
assertClose(logestStats[4][1], 0.00507376135958447, 1e-14);
assertClose(exponentialSheet.getRange("K1").values[0][0], 1.9854301969018, 1e-12);
assert.equal(exponentialSheet.getRange("L1").values[0][0], 1);
const exactGrowth = exponentialSheet.getRange("M4:O4").values[0];
assertClose(exactGrowth[0], 192);
assertClose(exactGrowth[1], 384);
assertClose(exactGrowth[2], 768);
for (const value of exponentialSheet.getRange("V1:V2").values.flat()) assertClose(value, 8);
assertClose(exponentialSheet.getRange("X1:X6").values[0][0], 6.196881018771227, 1e-12);
assertClose(exponentialSheet.getRange("Y1:Z1").values[0][0], 1.71968316255041, 1e-12);
assertClose(exponentialSheet.getRange("Y1:Z1").values[0][1], 3.6035015947826325, 1e-12);
assert.deepEqual(exponentialSheet.getRange("AA1:AA8").values.flat(), ["#VALUE!", "#N/A", "#VALUE!", "#VALUE!", "#VALUE!", "#VALUE!", "#NUM!", "#VALUE!"]);
assert.equal(exponentialSheet.getRange("AF1").values[0][0], "#NUM!");
assert.equal(exponentialSheet.getRange("AH1").values[0][0], "#SPILL!");
assert.deepEqual(exponentialSheet.store.get("AH1").spillError, { type: "blocked", addresses: ["AH2"] });
assert.equal(exponentialSheet.store.get("D1").spillRange, "D1:D3");
assert.equal(exponentialSheet.store.get("H1").spillRange, "H1:I5");
assert.equal(exponentialSheet.store.get("M4").spillRange, "M4:O4");
exponentialSheet.getRange("AH1:AH3").clear();

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
legacyArraySheet.getRange("I1:I3").values = [[8], [9], [10]];
legacyArraySheet.getRange("J1").formulas = [["=GROWTH(B1:B5,A1:A5,I1:I3)"]];
legacyArraySheet.store.get("J1").formulaType = "array";
legacyArraySheet.store.get("J1").arrayRef = "J1:J3";
legacyArraySheet.getRange("J1:J3").values = Array.from({ length: 3 }, () => [0]);
legacyArraySheet.store.get("J1").formula = "=GROWTH(B1:B5,A1:A5,I1:I3)";
legacyArraySheet.store.get("J1").formulaType = "array";
legacyArraySheet.store.get("J1").arrayRef = "J1:J3";
legacyArraySheet.getRange("L1").formulas = [["=LOGEST(B1:B5,A1:A5,TRUE(),TRUE())"]];
legacyArraySheet.store.get("L1").formulaType = "array";
legacyArraySheet.store.get("L1").arrayRef = "L1:M5";
legacyArraySheet.getRange("L1:M5").values = Array.from({ length: 5 }, () => [0, 0]);
legacyArraySheet.store.get("L1").formula = "=LOGEST(B1:B5,A1:A5,TRUE(),TRUE())";
legacyArraySheet.store.get("L1").formulaType = "array";
legacyArraySheet.store.get("L1").arrayRef = "L1:M5";
workbook.recalculate();
assert.deepEqual(legacyArraySheet.getRange("D1:D3").values, [[13], [15], [17]]);
assert.equal(legacyArraySheet.store.get("D1").spillRange, "D1:D3");
assert.deepEqual(legacyArraySheet.getRange("F1:G1").values, [[2, 1]]);
assert.equal(legacyArraySheet.store.get("F1").spillRange, "F1:G5");
assert.equal(legacyArraySheet.store.get("J1").spillRange, "J1:J3");
assert.equal(legacyArraySheet.store.get("L1").spillRange, "L1:M5");

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
mismatchedLegacySheet.getRange("F1:F3").values = [[8], [9], [10]];
mismatchedLegacySheet.getRange("G1:G2").values = [[160], [275]];
Object.assign(mismatchedLegacySheet.store.get("G1"), {
  formula: "=GROWTH(B1:B5,A1:A5,F1:F3)",
  formulaType: "array",
  arrayRef: "G1:G2",
});
mismatchedLegacySheet.getRange("I1:J1").values = [[2, 1]];
Object.assign(mismatchedLegacySheet.store.get("I1"), {
  formula: "=LOGEST(B1:B5,A1:A5,TRUE,TRUE)",
  formulaType: "array",
  arrayRef: "I1:J1",
});
mismatchedLegacyWorkbook.recalculate();
assert.equal(mismatchedLegacySheet.getRange("D1").values[0][0], "#SPILL!");
assert.equal(mismatchedLegacySheet.getRange("D2").values[0][0], 15);
assert.equal(mismatchedLegacySheet.store.get("D2").spillParent, undefined);
assert.equal(mismatchedLegacySheet.getRange("G1").values[0][0], "#SPILL!");
assert.equal(mismatchedLegacySheet.getRange("G2").values[0][0], 275);
assert.equal(mismatchedLegacySheet.store.get("G2").spillParent, undefined);
assert.equal(mismatchedLegacySheet.getRange("I1").values[0][0], "#SPILL!");
assert.equal(mismatchedLegacySheet.getRange("J1").values[0][0], 1);
assert.equal(mismatchedLegacySheet.store.get("J1").spillParent, undefined);

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
const importedExponentialSheet = imported.worksheets.getItem("Exponential forecast");
assert.deepEqual(importedExponentialSheet.getRange("D1:D3").values, exponentialSheet.getRange("D1:D3").values);
assert.equal(importedExponentialSheet.store.get("D1").formulaType, "dynamicArray");
assert.equal(importedExponentialSheet.store.get("D1").dynamicArrayRef, "D1:D3");
assert.deepEqual(importedExponentialSheet.getRange("H1:I5").values, logestStats);
assert.equal(importedExponentialSheet.store.get("H1").dynamicArrayRef, "H1:I5");
const importedLegacyArraySheet = imported.worksheets.getItem("Legacy array interop");
assert.equal(importedLegacyArraySheet.store.get("D1").formulaType, "array");
assert.equal(importedLegacyArraySheet.store.get("D1").arrayRef, "D1:D3");
assert.deepEqual(importedLegacyArraySheet.getRange("D1:D3").values, [[13], [15], [17]]);
assert.equal(importedLegacyArraySheet.store.get("F1").formulaType, "array");
assert.equal(importedLegacyArraySheet.store.get("F1").arrayRef, "F1:G5");
assert.equal(importedLegacyArraySheet.store.get("J1").arrayRef, "J1:J3");
assert.equal(importedLegacyArraySheet.store.get("L1").arrayRef, "L1:M5");
imported.recalculate();
assert.deepEqual(importedLegacyArraySheet.getRange("D1:D3").values, [[13], [15], [17]]);
assert.deepEqual(importedLegacyArraySheet.getRange("F1:G1").values, [[2, 1]]);
assert.equal(importedLegacyArraySheet.store.get("J1").spillRange, "J1:J3");
assert.equal(importedLegacyArraySheet.store.get("L1").spillRange, "L1:M5");
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
