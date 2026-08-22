import assert from "node:assert/strict";
import crypto from "node:crypto";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { FileBlob, Presentation, PresentationFile } from "../src/index.mjs";
import { runTemplateConditionedGeneration } from "../scripts/pptx-template-generation.mjs";

const workspace = await mkdtemp(path.join(os.tmpdir(), "office-kit-template-generation-fixture-"));
const assetsDir = path.join(workspace, "assets");
const sourcePath = path.join(assetsDir, "fixture.pptx");
const fixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
for (let index = 0; index < 4; index += 1) {
  const slide = fixture.slides.add({ name: `Fixture ${index + 1}` });
  slide.shapes.add({
    name: "Fixture title",
    position: { left: 40, top: 40, width: 520, height: 90 },
    text: [{ runs: [{ text: `Source title ${index + 1}`, style: { bold: true, fontSize: 26, fontFamily: "Arial" } }] }],
  });
}
const source = await PresentationFile.exportPptx(fixture);
await mkdir(assetsDir, { recursive: true });
await writeFile(sourcePath, source.bytes);
const sourceBytes = await readFile(sourcePath);
const sourceSha256 = crypto.createHash("sha256").update(sourceBytes).digest("hex");
const outputDir = path.join(workspace, "output");

try {
  const evidence = await runTemplateConditionedGeneration({
    assetsDir,
    outputDir,
    generatedSlides: 3,
    definitions: [{
      id: "authored-fixture",
      fileName: "fixture.pptx",
      sourceSha256,
      minimumGeneratedSlides: 3,
      content: ["Generated evidence", "Generated decision", "Generated next step"],
    }],
  });
  const result = evidence.sources[0];
  assert.equal(evidence.schema, "office-kit/pptx-template-conditioned-generation-evidence/v1");
  assert.equal(result.generatedSlides, 3);
  assert.equal(result.outputSlides, result.sourceSlides + 3);
  assert.equal(result.profile.schema, "office-kit/pptx-design-profile/v1");
  assert.equal(result.profile.source.sourceBound, true);
  assert.equal(result.verification.allTargetsRoundTrip, true);
  assert.equal(result.verification.noNewIssues, true);
  assert.equal(result.packageOracle.sourceProtected, true);
  assert.equal(result.packageOracle.nonTargetPartsPreserved, true);
  assert.deepEqual(result.packageOracle.missingParts, []);
  assert.ok(result.packageOracle.addedParts.some((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)));
  assert.ok(result.montageBytes > 0);

  // A generated deck is a durable source-bound revision, not a one-shot
  // export. Reopen it in a fresh process-shaped object and perform one local
  // conversational edit against the generated clone.
  const generatedBytes = await readFile(path.join(outputDir, "authored-fixture.pptx"));
  const resumed = await PresentationFile.importPptx(new FileBlob(generatedBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));
  const continuation = result.selection[0];
  const continuationSlide = resumed.slides.items[continuation.outputSlide - 1];
  const continuationShape = continuationSlide.shapes.items.find((shape) =>
    (!continuation.target.name || shape.name === continuation.target.name) &&
    (shape.text?.paragraphs || []).some((paragraph) => (paragraph.runs || []).some((run) => run.text === continuation.value)));
  assert.ok(continuationShape);
  continuationShape.text.replace(continuation.value, "Reviewed evidence");
  const continued = await PresentationFile.exportPptx(resumed);
  const continuedPresentation = await PresentationFile.importPptx(continued.bytes);
  assert.match(continuedPresentation.slides.items.flatMap((slide) => slide.shapes.items.map((shape) => shape.text?.value || "")).join("\n"), /Reviewed evidence/u);
  assert.match(continuedPresentation.slides.items[0].shapes.items[0].text.value, /Source title/u);
  console.log("pptx template-conditioned generation smoke ok");
} finally {
  await rm(workspace, { recursive: true, force: true });
}
