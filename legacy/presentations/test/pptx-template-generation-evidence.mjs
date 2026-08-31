import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const manifest = JSON.parse(await readFile(path.resolve(import.meta.dirname, "../evals/pptx-generation/template-conditioned.v1.json"), "utf8"));
const evidence = JSON.parse(await readFile(path.resolve(import.meta.dirname, "../evals/pptx-generation/evidence.v1.json"), "utf8"));

assert.equal(manifest.schema, "office-kit/pptx-template-conditioned-generation/v1");
assert.equal(manifest.generatedSlides, 10);
assert.equal(manifest.sourcePolicy, "external-hash-only");
assert.equal(manifest.sources.length, 3);
for (const source of manifest.sources) {
  assert.match(source.sha256, /^[a-f0-9]{64}$/u);
  assert.ok(Number.isInteger(source.sourceSlides) && source.sourceSlides > 0);
}
assert.deepEqual(manifest.acceptance, {
  allTargetsRoundTrip: true,
  sourceProtected: true,
  nonTargetPartsPreserved: true,
  logicalSourceSlidesPreserved: true,
  noNewReviewIssueCategory: true,
  visualReview: "record-available-or-unavailable",
  windowsPowerPoint: "separate-host-lane",
});
assert.equal(evidence.schema, "office-kit/pptx-template-conditioned-generation-evidence/v1");
assert.equal(evidence.runs.length, manifest.sources.length);
for (const run of evidence.runs) {
  assert.match(run.sourceSha256, /^[a-f0-9]{64}$/u);
  assert.match(run.outputSha256, /^[a-f0-9]{64}$/u);
  assert.equal(run.selectedTargets, manifest.generatedSlides);
  assert.equal(run.allTargetsRoundTrip, true);
  assert.equal(run.noNewReviewIssueCategory, true);
  assert.equal(run.nonTargetPartsPreserved, true);
  assert.equal(run.logicalSourceSlidesPreserved, true);
}
console.log("pptx template-conditioned evidence smoke ok");
