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
sheet.getRange("H1:H26").formulas = [
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
];

const results = sheet.getRange("H1:H26").values.flat();
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

const spillSheet = workbook.worksheets.add("Spill statistics");
spillSheet.getRange("A1").formulas = [["=SEQUENCE(5)"]];
spillSheet.getRange("C1").formulas = [["=SEQUENCE(5,1,2,2)"]];
spillSheet.getRange("E1:E4").formulas = [
  ["=CORREL(A1#,C1#)"],
  ["=COVARIANCE.P(A1#,C1#)"],
  ["=VAR.P(A1#)"],
  ["=STDEV.S(C1#)"],
];
const spillResults = spillSheet.getRange("E1:E4").values.flat();
assertClose(spillResults[0], 1);
assertClose(spillResults[1], 4);
assertClose(spillResults[2], 2);
assertClose(spillResults[3], Math.sqrt(10));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
const imported = await SpreadsheetFile.importXlsx(xlsx);
assert.deepEqual(imported.worksheets.getItem("Statistics").getRange("H1:H26").formulas, sheet.getRange("H1:H26").formulas);
assert.deepEqual(imported.worksheets.getItem("Statistics").getRange("H1:H26").values, sheet.getRange("H1:H26").values);
assert.deepEqual(imported.worksheets.getItem("Spill statistics").getRange("E1:E4").formulas, spillSheet.getRange("E1:E4").formulas);
assert.deepEqual(imported.worksheets.getItem("Spill statistics").getRange("E1:E4").values, spillSheet.getRange("E1:E4").values);

console.log("spreadsheet statistical formula tests passed");
