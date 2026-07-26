import os from "node:os";
import path from "node:path";

import { DocumentFile, DocumentModel } from "office-kit";
import { nativeOfficeStatus, renderFileWithNativeOffice } from "office-kit/native/office-bridge";

const outputDir = process.env.OUTPUT_DIR || path.join(os.tmpdir(), "office-kit-artifact-examples");

const status = await nativeOfficeStatus({ timeoutMs: 10_000 });
if (!status.available) {
  console.log(`skipped native Office render example: ${status.error?.message || "bridge unavailable"}`);
  process.exit(0);
}

const document = DocumentModel.create({ paragraphs: ["Native Office bridge render", "This example expects a JSON stdin/stdout bridge command."] });
const docx = await DocumentFile.exportDocx(document);
const pdf = await renderFileWithNativeOffice(docx, {
  artifactKind: "document",
  inputType: docx.type,
  format: "pdf",
  outputType: "application/pdf",
  timeoutMs: 60_000,
});
await pdf.save(path.join(outputDir, "native-office-render.pdf"));
console.log(`saved ${path.join(outputDir, "native-office-render.pdf")}`);
