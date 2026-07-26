import os from "node:os";
import path from "node:path";

import { FileBlob, PdfArtifact, PdfFile } from "office-kit";

const outputDir = process.env.OUTPUT_DIR || path.join(os.tmpdir(), "office-kit-artifact-examples");

const pdf = PdfArtifact.create({
  pages: [
    {
      text: "PDF artifact report\nGenerated with clean-room metadata",
      tables: [{ name: "pdf-metrics", values: [["Metric", "Value"], ["Pages", "1"]], bbox: [72, 160, 360, 72] }],
      images: [{ name: "pdf-image", alt: "Modeled image placeholder", prompt: "placeholder", bbox: [72, 260, 180, 100] }],
    },
  ],
});
const pdfFile = await PdfFile.exportPdf(pdf);
const pdfPath = path.join(outputDir, "modeled-report.pdf");
await pdfFile.save(pdfPath);
const loaded = await PdfFile.importPdf(await FileBlob.load(pdfPath));
const svg = await loaded.render({ pageIndex: 0 });
await svg.save(path.join(outputDir, "modeled-report.svg"));
console.log(loaded.inspect({ kind: "page,text,table,image", maxChars: 5000 }).ndjson);
console.log(`saved ${pdfPath}`);
console.log(`saved ${path.join(outputDir, "modeled-report.svg")}`);
