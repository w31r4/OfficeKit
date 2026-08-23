import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const root = path.resolve(import.meta.dirname, "..");
const evidence = JSON.parse(await readFile(path.join(root, "evals/pptx-lossless/svg-leaves.v1.json"), "utf8"));
const manifest = JSON.parse(await readFile(path.join(root, "evals/pptx-lossless/manifest.v1.json"), "utf8"));

assert.equal(evidence.schema, "office-kit/pptx-svg-leaf-evidence/v1");
assert.equal(evidence.editSourceId, "mckinsey-customer-loyalty");
assert.equal(evidence.repeats, 3);
assert.deepEqual(evidence.sources.map((source) => source.id), [
  "suanzhi-future-2026",
  "blue-gray-acid-template",
  "mckinsey-customer-loyalty",
]);

for (const source of evidence.sources) {
  const declared = manifest.sources.find((candidate) => candidate.id === source.id);
  assert.ok(declared, `missing frozen manifest source ${source.id}`);
  assert.equal(source.sourceSha256, declared.sha256);
  assert.equal(source.slideCount, declared.inventory.slideCount);
  assert.equal(source.revisionBoundLeafCount, source.svgLeafCount);
  if (source.id === "mckinsey-customer-loyalty") {
    assert.equal(source.imageCount, 8);
    assert.equal(source.svgImageCount, 8);
    assert.equal(source.supportedSvgImageCount, 8);
    assert.equal(source.svgLeafCount, 487);
    assert.deepEqual(source.leafKinds, {
      svgFillRgb: 373,
      svgOpacity: 35,
      svgStrokeRgb: 58,
      svgTransformScalar: 21,
    });
  } else {
    assert.equal(source.svgImageCount, 0);
    assert.equal(source.supportedSvgImageCount, 0);
    assert.equal(source.svgLeafCount, 0);
    assert.deepEqual(source.leafKinds, {});
  }
}

assert.equal(evidence.runs.length, 3);
assert.equal(new Set(evidence.runs.map((run) => run.outputArchiveSha256)).size, 1);
assert.equal(new Set(evidence.runs.map((run) => run.canonicalOutputSha256)).size, 1);
assert.equal(new Set(evidence.runs.map((run) => run.editedSvgSha256)).size, 1);
assert.equal(new Set(evidence.runs.map((run) => JSON.stringify({
  edits: run.edits,
  changedExistingParts: run.changedExistingParts,
  deletedParts: run.deletedParts,
  addedParts: run.addedParts,
}))).size, 1);

const source = evidence.sources.find((candidate) => candidate.id === evidence.editSourceId);
for (const [index, run] of evidence.runs.entries()) {
  assert.equal(run.repeat, index + 1);
  assert.equal(run.sourceSha256, source.sourceSha256);
  assert.equal(run.sourceSvgSha256, "4c8ca5ad52f261b7f1b1466232a29e793c6cdee98631d5b231f75582308241ed");
  assert.equal(run.editedSvgSha256, "32450ce15b2d80c63e9490ae8920c517f984a9041c38e05dd15cf3184ab7b299");
  assert.match(run.outputArchiveSha256, /^[0-9a-f]{64}$/u);
  assert.match(run.canonicalOutputSha256, /^[0-9a-f]{64}$/u);
  assert.equal(run.editedSlide, 3);
  assert.deepEqual(run.edits.map((edit) => edit.leafKind), [
    "svgFillRgb",
    "svgStrokeRgb",
    "svgOpacity",
    "svgTransformScalar",
  ]);
  assert.deepEqual(run.edits.map((edit) => edit.value), ["#0F766E", "#DC2626", 0.65, -85]);
  assert.equal(run.edits.every((edit) => edit.sourceRevisionSha256 === source.sourceSha256), true);
  assert.equal(run.edits[0].mutation.beforeSha256, run.sourceSvgSha256);
  for (let editIndex = 1; editIndex < run.edits.length; editIndex += 1) {
    assert.equal(run.edits[editIndex].mutation.beforeSha256, run.edits[editIndex - 1].mutation.afterSha256);
  }
  assert.equal(run.edits.at(-1).mutation.afterSha256, run.editedSvgSha256);
  assert.equal(run.edits.every((edit) => Number.isSafeInteger(edit.mutation.offset) && edit.mutation.offset >= 0), true);
  assert.deepEqual(run.changedExistingParts, [
    "ppt/slides/_rels/slide3.xml.rels",
    "ppt/slides/slide3.xml",
  ]);
  assert.deepEqual(run.deletedParts, []);
  assert.deepEqual(run.addedParts, ["ppt/media/image9.svg"]);
  assert.equal(run.replacementSvgPart, "ppt/media/image9.svg");
  assert.equal(run.replacementSvgBytesMatch, true);
  assert.equal(run.reimported, true);
  assert.ok(run.classifiedOperations.includes("svg-style"));
  assert.ok(run.classifiedOperations.includes("svg-transform"));
}

console.log("PPTX SVG leaf benchmark evidence ok");
