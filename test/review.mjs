import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  DocumentFile,
  DocumentModel,
  PdfArtifact,
  Presentation,
  reviewArtifact,
  Workbook,
} from "../src/index.mjs";

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "officekit-review-"));
try {
  const document = DocumentModel.create({ paragraphs: ["Quarterly review document"] });

  const workbook = Workbook.create();
  workbook.worksheets.add("Data").getRange("A1:B2").values = [
    ["Metric", "Value"],
    ["Revenue", 42],
  ];

  const presentation = Presentation.create();
  presentation.slides.add({ name: "Review" }).shapes.add({
    geometry: "textbox",
    text: "Quarterly review presentation",
    position: { left: 20, top: 20, width: 400, height: 80 },
  });

  const pdf = PdfArtifact.create({ pages: [{ text: "Quarterly review PDF" }] });

  for (const [format, artifact, expectedText] of [
    ["docx", document, "Quarterly review document"],
    ["xlsx", workbook, "Revenue"],
    ["pptx", presentation, "Quarterly review presentation"],
    ["pdf", pdf, "Quarterly review PDF"],
  ]) {
    const report = await reviewArtifact(artifact, {
      outputPath: path.join(temporary, `review.${format}`),
      contentView: "anydoc",
      visualReview: "complete",
    });
    assert.equal(report.schemaVersion, 1);
    assert.equal(report.format, format);
    assert.equal(report.verdict, "passed", JSON.stringify(report, null, 2));
    assert.equal(report.semantic.status, "passed");
    assert.equal(report.structural.status, "passed");
    assert.equal(report.layout.status, "passed");
    assert.equal(report.contentView.status, "ready");
    assert.equal(report.contentView.provider, "anydoc");
    assert.equal(report.contentView.providerVersion, "0.1.3");
    assert.equal(report.contentView.sourceSha256, report.delivery.sha256);
    assert.match(report.contentView.markdown, new RegExp(expectedText, "i"));
    for (const step of ["6. Semantic", "7. Structural", "8. Layout", "9. AnyDoc", "10. Visual", "11. Delivery"]) {
      assert.match(report.summary.markdown, new RegExp(step.replace(".", "\\.")));
    }
  }

  const unavailableVisualReview = await reviewArtifact(document, {
    outputPath: path.join(temporary, "unavailable-visual.docx"),
  });
  assert.equal(unavailableVisualReview.verdict, "passed-with-limitations");
  assert.equal(unavailableVisualReview.visualReview, "unavailable");

  const disabledContentView = await reviewArtifact(document, {
    outputPath: path.join(temporary, "disabled-content.docx"),
    contentView: false,
    visualReview: "complete",
  });
  assert.equal(disabledContentView.contentView.status, "not-requested");
  assert.equal(disabledContentView.verdict, "passed");

  const longDocument = DocumentModel.create({ paragraphs: [`Review ${"content ".repeat(500)}`] });
  const truncated = await reviewArtifact(longDocument, {
    outputPath: path.join(temporary, "truncated.docx"),
    contentView: "anydoc",
    maxContentChars: 240,
    maxSummaryChars: 3_000,
    visualReview: "complete",
  });
  assert.equal(truncated.contentView.truncated, true);
  assert.ok(truncated.contentView.originalChars > truncated.contentView.chars);
  assert.match(truncated.contentView.markdown, /content view truncated/);
  assert.match(truncated.summary.markdown, /11\. Delivery review/);

  const savedPath = path.join(temporary, "source.docx");
  await (await DocumentFile.exportDocx(document)).save(savedPath);
  const collision = await reviewArtifact(savedPath, {
    source: savedPath,
    visualReview: "complete",
  });
  assert.equal(collision.delivery.ok, false);
  assert.equal(collision.verdict, "failed");
  assert.ok(collision.delivery.issues.some((issue) => issue.type === "inputOutputCollision"));

  await assert.rejects(
    reviewArtifact(new Uint8Array(20), { format: "pdf", maxBytes: 10 }),
    /exceeds maxBytes/,
  );
  await assert.rejects(
    reviewArtifact(document, { maxBytes: Number.NaN }),
    /maxBytes must be a positive safe integer/,
  );
  await assert.rejects(
    reviewArtifact(document, { visualReview: "ocr-complete" }),
    /visualReview must be complete, unavailable, or requires-human/,
  );

  const corrupt = await reviewArtifact(
    new Blob(["not an OOXML ZIP"], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }),
    { outputPath: path.join(temporary, "corrupt.docx"), contentView: "anydoc", visualReview: "complete" },
  );
  assert.equal(corrupt.verdict, "failed");
  assert.equal(corrupt.structural.status, "failed");
  assert.equal(corrupt.contentView.status, "blocked");

  const rootProbe = spawnSync(process.execPath, [
    "--input-type=module",
    "-e",
    "import Module from 'node:module'; await import('./src/index.mjs'); console.log(JSON.stringify(Object.keys(Module._cache).filter((name) => /firecrawl|anydoc/i.test(name))))",
  ], {
    cwd: path.resolve(import.meta.dirname, ".."),
    encoding: "utf8",
    env: process.env,
  });
  assert.equal(rootProbe.status, 0, rootProbe.stderr);
  assert.deepEqual(JSON.parse(rootProbe.stdout), [], "root import must not initialize AnyDoc or its native binding");
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("post-edit review smoke ok");
