import fs from "node:fs/promises";

import {
  DocumentFile,
  DocumentModel,
  PdfArtifact,
  PdfFile,
  reviewArtifact,
  SpreadsheetFile,
  Workbook,
} from "office-kit";

const publicSpecifiers = [
  "office-kit",
  "office-kit/live",
  "office-kit/live/protocol",
  "office-kit/live/adapters/powerpoint",
  "office-kit/powerpoint-live",
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
const documentReview = await reviewArtifact("standalone.docx", {
  contentView: "anydoc",
  layout: false,
  visualReview: "unavailable",
});
if (
  documentReview.contentView.status !== "ready" ||
  documentReview.contentView.providerVersion !== "0.1.3" ||
  !documentReview.contentView.markdown.includes("standalone DOCX")
) {
  process.exit(16);
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
  "standalone.pdf",
]) {
  if ((await fs.stat(filename)).size < 100) process.exit(15);
}

console.log(
  JSON.stringify({
    argv: process.argv.slice(2),
    cwd: process.cwd(),
    publicSubpaths: publicSpecifiers.length,
    anydoc: documentReview.contentView.status,
  }),
);
