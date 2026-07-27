import fs from "node:fs/promises";

import {
  DocumentFile,
  DocumentModel,
  PdfArtifact,
  PdfFile,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "office-kit";

const publicSpecifiers = [
  "office-kit",
  "office-kit/presentation-jsx",
  "office-kit/presentation-jsx/jsx-runtime",
  "office-kit/presentation-jsx/jsx-dev-runtime",
  "office-kit/renderers/playwright",
  "office-kit/renderers/sharp",
  "office-kit/renderers/canvas",
  "office-kit/renderers/poppler",
  "office-kit/renderers/libreoffice",
  "office-kit/pdf/pdfjs",
  "office-kit/pdf/mupdf",
  "office-kit/pdf/providers",
  "office-kit/native/office-bridge",
  "office-kit/codec",
  "office-kit/codec/wire",
];
for (const specifier of publicSpecifiers) await import(specifier);

const document = DocumentModel.create({ paragraphs: ["standalone DOCX"] });
const docx = await DocumentFile.exportDocx(document);
await docx.save("standalone.docx");
if ((await DocumentFile.importDocx(docx)).blocks[0].text !== "standalone DOCX") {
  process.exit(11);
}

const workbook = Workbook.create();
workbook.worksheets
  .add("Data")
  .getRange("A1:B2").values = [
    ["Label", "Value"],
    ["standalone XLSX", 7],
  ];
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save("standalone.xlsx");
if (
  (await SpreadsheetFile.importXlsx(xlsx))
    .worksheets.getItem("Data")
    .getRange("B2").values[0][0] !== 7
) {
  process.exit(12);
}

const presentation = Presentation.create();
presentation.slides.add({ name: "Standalone" }).shapes.add({
  geometry: "textbox",
  text: "standalone PPTX",
  position: { left: 40, top: 40, width: 400, height: 80 },
});
const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save("standalone.pptx");
if ((await PresentationFile.importPptx(pptx)).slides.count !== 1) {
  process.exit(13);
}

const pdf = await PdfFile.exportPdf(
  PdfArtifact.create({ pages: [{ text: "standalone PDF" }] }),
);
await pdf.save("standalone.pdf");
if (!(await PdfFile.importPdf(pdf)).extractText().includes("standalone PDF")) {
  process.exit(14);
}

for (const filename of [
  "standalone.docx",
  "standalone.xlsx",
  "standalone.pptx",
  "standalone.pdf",
]) {
  if ((await fs.stat(filename)).size < 100) process.exit(15);
}

console.log(
  JSON.stringify({
    argv: process.argv.slice(2),
    cwd: process.cwd(),
    publicSubpaths: publicSpecifiers.length,
  }),
);
