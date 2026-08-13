import assert from "node:assert/strict";

import JSZip from "jszip";

import { SpreadsheetFile, Workbook } from "../src/index.mjs";
import {
  modelFormulaFromXlsx,
  xlsxFormulaFromModel,
} from "../src/codecs/office-kit-spreadsheet-formula-syntax.mjs";

assert.equal(
  xlsxFormulaFromModel("=STDEV.S(A1:A3)+IFNA(XMATCH(4,A1:A3),0)"),
  "=_xlfn.STDEV.S(A1:A3)+_xlfn.IFNA(_xlfn.XMATCH(4,A1:A3),0)",
);
assert.equal(
  xlsxFormulaFromModel('="STDEV.S(A1:A3)"&STDEV.S(A1:A3)'),
  '="STDEV.S(A1:A3)"&_xlfn.STDEV.S(A1:A3)',
);
assert.equal(
  xlsxFormulaFromModel("=XLOOKUP(A1,Table1[XMATCH],Table1[Result])"),
  "=_xlfn.XLOOKUP(A1,Table1[XMATCH],Table1[Result])",
);
assert.equal(
  xlsxFormulaFromModel("=Table1[Column']XMATCH(]+XMATCH(A1,A1:A3)"),
  "=Table1[Column']XMATCH(]+_xlfn.XMATCH(A1,A1:A3)",
);
assert.equal(
  xlsxFormulaFromModel("='STDEV.S('!A1+STDEV.S(A1:A3)"),
  "='STDEV.S('!A1+_xlfn.STDEV.S(A1:A3)",
);
assert.equal(
  xlsxFormulaFromModel("=FILTER('Source Data'!A1:A3,'Source Data'!A1:A3>1)"),
  "=_xlfn._xlws.FILTER('Source Data'!A1:A3,'Source Data'!A1:A3>1)",
);
assert.equal(
  xlsxFormulaFromModel("=SUM('Source Data'!$A$1#)"),
  "=SUM(_xlfn.ANCHORARRAY('Source Data'!$A$1))",
);
assert.equal(xlsxFormulaFromModel("=MYRANK.AVG(A1)"), "=MYRANK.AVG(A1)");
assert.equal(xlsxFormulaFromModel("=MYA1#"), "=_xlfn.ANCHORARRAY(MYA1)");
assert.equal(xlsxFormulaFromModel("=NamedRange1#"), "=NamedRange1#");
assert.equal(xlsxFormulaFromModel("=A1#suffix"), "=A1#suffix");
assert.equal(xlsxFormulaFromModel("=XFE1#+A1048577#"), "=XFE1#+A1048577#");
assert.equal(xlsxFormulaFromModel("=_xlfn.RANK.AVG(A1,A1:A3)"), "=_xlfn.RANK.AVG(A1,A1:A3)");
assert.equal(
  modelFormulaFromXlsx("=_xlfn.STDEV.S(A1:A3)+_xlfn.IFNA(_xlfn.XMATCH(4,A1:A3),0)"),
  "=STDEV.S(A1:A3)+IFNA(XMATCH(4,A1:A3),0)",
);
assert.equal(modelFormulaFromXlsx("=SUM(_xlfn.ANCHORARRAY('Source Data'!$A$1))"), "=SUM('Source Data'!$A$1#)");
assert.equal(modelFormulaFromXlsx('="_xlfn.STDEV.S("&_xlfn.STDEV.S(A1:A3)'), '="_xlfn.STDEV.S("&STDEV.S(A1:A3)');

async function worksheetXml(file) {
  const zip = await JSZip.loadAsync(new Uint8Array(await file.arrayBuffer()));
  return zip.file("xl/worksheets/sheet1.xml").async("text");
}

const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
const sheet = workbook.worksheets.add("Formula syntax");
sheet.getRange("A1:A3").values = [[1], [2], [4]];
sheet.getRange("C1:C3").formulas = [
  ["=STDEV.S(A1:A3)"],
  ["=IFNA(XMATCH(4,A1:A3),0)"],
  ['="STDEV.S("&STDEV.S(A1:A3)'],
];
sheet.getRange("E1").formulas = [["=SEQUENCE(3)"]];
sheet.getRange("F1").formulas = [["=SUM(E1#)"]];
sheet.getRange("H1").formulas = [["=FILTER(A1:A3,A1:A3>1)"]];
sheet.getRange("J1:K4").values = [["Value", "Rank"], [1, 1], [2, 2], [4, 3]];
sheet.tables.add({
  name: "FutureFormulaTable",
  range: "J1:K4",
  columnDefinitions: [
    { name: "Value" },
    { name: "Rank", calculatedColumnFormula: "=RANK.EQ([@Value],[Value])" },
  ],
});
workbook.recalculate();

const sourceFree = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
const sourceFreeXml = await worksheetXml(sourceFree);
assert.match(sourceFreeXml, /_xlfn\.STDEV\.S\(A1:A3\)/);
assert.match(sourceFreeXml, /_xlfn\.IFNA\(_xlfn\.XMATCH\(4,A1:A3\),0\)/);
assert.match(sourceFreeXml, /"STDEV\.S\("&amp;_xlfn\.STDEV\.S\(A1:A3\)/);
assert.match(sourceFreeXml, /_xlfn\.SEQUENCE\(3\)/);
assert.match(sourceFreeXml, /SUM\(_xlfn\.ANCHORARRAY\(E1\)\)/);
assert.match(sourceFreeXml, /_xlfn\._xlws\.FILTER\(A1:A3,A1:A3&gt;1\)/);
const sourceFreeZip = await JSZip.loadAsync(new Uint8Array(await sourceFree.arrayBuffer()));
const tableXml = await sourceFreeZip.file("xl/tables/table1.xml").async("text");
assert.match(tableXml, /_xlfn\.RANK\.EQ\(\[@Value\],\[Value\]\)/);

const imported = await SpreadsheetFile.importXlsx(sourceFree);
const importedSheet = imported.worksheets.getItem("Formula syntax");
assert.deepEqual(importedSheet.getRange("C1:C3").formulas.flat(), [
  "=STDEV.S(A1:A3)",
  "=IFNA(XMATCH(4,A1:A3),0)",
  '="STDEV.S("&STDEV.S(A1:A3)',
]);
assert.equal(importedSheet.getRange("E1").formulas[0][0], "=SEQUENCE(3)");
assert.equal(importedSheet.getRange("F1").formulas[0][0], "=SUM(E1#)");
assert.equal(importedSheet.getRange("H1").formulas[0][0], "=FILTER(A1:A3,A1:A3>1)");
assert.equal(importedSheet.tables.items.find((table) => table.name === "FutureFormulaTable").columnDefinitions[1].calculatedColumnFormula, "=RANK.EQ([@Value],[Value])");

const rawSourceZip = await JSZip.loadAsync(new Uint8Array(await sourceFree.arrayBuffer()));
rawSourceZip.file("xl/worksheets/sheet1.xml", sourceFreeXml.replace("_xlfn.STDEV.S(A1:A3)", "STDEV.S(A1:A3)"));
const rawSourceBytes = await rawSourceZip.generateAsync({ type: "uint8array" });
const rawImported = await SpreadsheetFile.importXlsx(rawSourceBytes);
const rawSheet = rawImported.worksheets.getItem("Formula syntax");
assert.equal(rawSheet.getRange("C1").formulas[0][0], "=STDEV.S(A1:A3)");
const rawPreserved = await SpreadsheetFile.exportXlsx(rawImported, { recalculate: false });
assert.match(await worksheetXml(rawPreserved), /<x:f>STDEV\.S\(A1:A3\)<\/x:f>/);

rawSheet.getRange("C1").formulas = [["=STDEV.P(A1:A3)"]];
rawImported.recalculate();
const edited = await SpreadsheetFile.exportXlsx(rawImported, { recalculate: false });
assert.match(await worksheetXml(edited), /<x:f>_xlfn\.STDEV\.P\(A1:A3\)<\/x:f>/);
const editedImported = await SpreadsheetFile.importXlsx(edited);
assert.equal(editedImported.worksheets.getItem("Formula syntax").getRange("C1").formulas[0][0], "=STDEV.P(A1:A3)");

console.log("spreadsheet formula package syntax tests passed");
