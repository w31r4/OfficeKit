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
  xlsxFormulaFromModel("=LET(rate,0.1,principal,1000,principal*(1+rate))"),
  "=_xlfn.LET(_xlpm.rate,0.1,_xlpm.principal,1000,_xlpm.principal*(1+_xlpm.rate))",
);
assert.equal(
  xlsxFormulaFromModel("=LET(x,2,LET(x,3,x)+x)"),
  "=_xlfn.LET(_xlpm.x,2,_xlfn.LET(_xlpm.x,3,_xlpm.x)+_xlpm.x)",
);
assert.equal(
  xlsxFormulaFromModel("=LET(x,y,y,2,x+y)"),
  "=_xlfn.LET(_xlpm.x,y,_xlpm.y,2,_xlpm.x+_xlpm.y)",
);
assert.equal(
  xlsxFormulaFromModel("=LET(SUM,2,SUM(A1:A2)+SUM)"),
  "=_xlfn.LET(_xlpm.SUM,2,SUM(A1:A2)+_xlpm.SUM)",
);
assert.equal(
  xlsxFormulaFromModel("=LET(E,2,1E+3+E)"),
  "=_xlfn.LET(_xlpm.E,2,1E+3+_xlpm.E)",
);
assert.equal(
  xlsxFormulaFromModel("=LET(A,2,IFERROR(#N/A,A))"),
  "=_xlfn.LET(_xlpm.A,2,IFERROR(#N/A,_xlpm.A))",
);
assert.equal(
  xlsxFormulaFromModel('=LET(label,"x,LET(y,1,y)",label&Table1[label])'),
  '=_xlfn.LET(_xlpm.label,"x,LET(y,1,y)",_xlpm.label&Table1[label])',
);
assert.equal(
  xlsxFormulaFromModel("=LET(x,2,Sheet1!x+'Source Data'!x+x)"),
  "=_xlfn.LET(_xlpm.x,2,Sheet1!x+'Source Data'!x+_xlpm.x)",
);
assert.equal(xlsxFormulaFromModel("=LET(x,1)"), "=LET(x,1)");
assert.equal(xlsxFormulaFromModel("=LET(A1,1,A1)"), "=LET(A1,1,A1)");
assert.equal(xlsxFormulaFromModel("=LET(XFE1,1,XFE1)"), "=LET(XFE1,1,XFE1)");
assert.equal(xlsxFormulaFromModel('=LET(x,1,"unterminated)'), '=LET(x,1,"unterminated)');
const seventeenBindings = `=LET(${Array.from({ length: 17 }, (_, index) => `n${index + 1},${index + 1}`).join(",")},n17)`;
assert.equal(xlsxFormulaFromModel(seventeenBindings), seventeenBindings);
const overNestedBudget = `=${"IF(TRUE,".repeat(65)}LET(x,1,x)${",0)".repeat(65)}`;
assert.equal(xlsxFormulaFromModel(overNestedBudget), overNestedBudget);
assert.equal(
  modelFormulaFromXlsx("=_xlfn.STDEV.S(A1:A3)+_xlfn.IFNA(_xlfn.XMATCH(4,A1:A3),0)"),
  "=STDEV.S(A1:A3)+IFNA(XMATCH(4,A1:A3),0)",
);
assert.equal(modelFormulaFromXlsx("=SUM(_xlfn.ANCHORARRAY('Source Data'!$A$1))"), "=SUM('Source Data'!$A$1#)");
assert.equal(modelFormulaFromXlsx('="_xlfn.STDEV.S("&_xlfn.STDEV.S(A1:A3)'), '="_xlfn.STDEV.S("&STDEV.S(A1:A3)');
assert.equal(
  modelFormulaFromXlsx("=_xlfn.LET(_xlpm.rate,0.1,_xlpm.principal,1000,_xlpm.principal*(1+_xlpm.rate))"),
  "=LET(rate,0.1,principal,1000,principal*(1+rate))",
);
assert.equal(
  modelFormulaFromXlsx("=_xlfn.LET(_xlpm.x,2,_xlfn.LET(_xlpm.x,3,_xlpm.x)+_xlpm.x)"),
  "=LET(x,2,LET(x,3,x)+x)",
);
assert.equal(modelFormulaFromXlsx("=_xlpm.orphan+1"), "=_xlpm.orphan+1");
assert.equal(
  modelFormulaFromXlsx("=_xlfn.LET(_xlpm.x,1,_xlpm.orphan+_xlpm.x)"),
  "=_xlfn.LET(_xlpm.x,1,_xlpm.orphan+_xlpm.x)",
);
assert.equal(
  modelFormulaFromXlsx("=_xlfn.LET(_xlpm.x,1,x)"),
  "=_xlfn.LET(_xlpm.x,1,x)",
);
assert.equal(
  modelFormulaFromXlsx("=_xlfn.LET(x,1,x)"),
  "=_xlfn.LET(x,1,x)",
);
assert.equal(modelFormulaFromXlsx("=LET(x,1,x)"), "=LET(x,1,x)");
assert.equal(
  modelFormulaFromXlsx("=_xlfn.LET(_xlpm.x,1,LET(y,2,y)+_xlpm.x)"),
  "=_xlfn.LET(_xlpm.x,1,LET(y,2,y)+_xlpm.x)",
);
assert.equal(
  modelFormulaFromXlsx("=LET(x,1,_xlfn.LET(_xlpm.y,2,_xlpm.y)+x)"),
  "=LET(x,1,_xlfn.LET(_xlpm.y,2,_xlpm.y)+x)",
);

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
sheet.getRange("I1").formulas = [["=LET(rate,0.1,principal,1000,principal*(1+rate))"]];
sheet.getRange("J1:L4").values = [["Value", "Rank", "Double"], [1, 1, 2], [2, 2, 4], [4, 3, 8]];
sheet.tables.add({
  name: "FutureFormulaTable",
  range: "J1:L4",
  columnDefinitions: [
    { name: "Value" },
    { name: "Rank", calculatedColumnFormula: "=RANK.EQ([@Value],[Value])" },
    { name: "Double", calculatedColumnFormula: "=LET(value,[@Value],value*2)" },
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
assert.match(sourceFreeXml, /_xlfn\.LET\(_xlpm\.rate,0\.1,_xlpm\.principal,1000,_xlpm\.principal\*\(1\+_xlpm\.rate\)\)/);
const sourceFreeZip = await JSZip.loadAsync(new Uint8Array(await sourceFree.arrayBuffer()));
const tableXml = await sourceFreeZip.file("xl/tables/table1.xml").async("text");
assert.match(tableXml, /_xlfn\.RANK\.EQ\(\[@Value\],\[Value\]\)/);
assert.match(tableXml, /_xlfn\.LET\(_xlpm\.value,\[@Value\],_xlpm\.value\*2\)/);

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
assert.equal(importedSheet.getRange("I1").formulas[0][0], "=LET(rate,0.1,principal,1000,principal*(1+rate))");
assert.equal(importedSheet.tables.items.find((table) => table.name === "FutureFormulaTable").columnDefinitions[1].calculatedColumnFormula, "=RANK.EQ([@Value],[Value])");
assert.equal(importedSheet.tables.items.find((table) => table.name === "FutureFormulaTable").columnDefinitions[2].calculatedColumnFormula, "=LET(value,[@Value],value*2)");

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

const rawLetSourceZip = await JSZip.loadAsync(new Uint8Array(await sourceFree.arrayBuffer()));
const rawLetXml = sourceFreeXml.replace(
  "_xlfn.LET(_xlpm.rate,0.1,_xlpm.principal,1000,_xlpm.principal*(1+_xlpm.rate))",
  "LET(rate,0.1,principal,1000,principal*(1+rate))",
);
rawLetSourceZip.file("xl/worksheets/sheet1.xml", rawLetXml);
const rawLetImported = await SpreadsheetFile.importXlsx(await rawLetSourceZip.generateAsync({ type: "uint8array" }));
assert.equal(rawLetImported.worksheets.getItem("Formula syntax").getRange("I1").formulas[0][0], "=LET(rate,0.1,principal,1000,principal*(1+rate))");
const rawLetPreserved = await SpreadsheetFile.exportXlsx(rawLetImported, { recalculate: false });
assert.match(await worksheetXml(rawLetPreserved), /<x:f>LET\(rate,0\.1,principal,1000,principal\*\(1\+rate\)\)<\/x:f>/);
rawLetImported.worksheets.getItem("Formula syntax").getRange("I1").formulas = [["=LET(rate,0.2,principal,1000,principal*(1+rate))"]];
rawLetImported.recalculate();
const editedLet = await SpreadsheetFile.exportXlsx(rawLetImported, { recalculate: false });
assert.match(await worksheetXml(editedLet), /<x:f>_xlfn\.LET\(_xlpm\.rate,0\.2,_xlpm\.principal,1000,_xlpm\.principal\*\(1\+_xlpm\.rate\)\)<\/x:f>/);
const editedLetImported = await SpreadsheetFile.importXlsx(editedLet);
assert.equal(editedLetImported.worksheets.getItem("Formula syntax").getRange("I1").formulas[0][0], "=LET(rate,0.2,principal,1000,principal*(1+rate))");

console.log("spreadsheet formula package syntax tests passed");
