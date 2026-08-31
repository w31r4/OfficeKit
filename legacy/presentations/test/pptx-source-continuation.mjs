import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const evidence = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/source-continuation.v2.json"), "utf8"));
assert.equal(evidence.schema, "office-kit/pptx-source-continuation-evidence/v2");
assert.deepEqual(evidence.sources.map((source) => source.id), [
  "suanzhi-future-2026",
  "blue-gray-acid-template",
  "mckinsey-customer-loyalty",
]);
for (const source of evidence.sources) {
  assert.equal(source.repetition, 1);
  assert.equal(source.repetitions, 3);
  assert.equal(source.deterministic, true);
  assert.equal(source.kind, "bounded-overlay");
  assert.equal(source.sourceSlideUnchanged, true);
  assert.equal(source.outputSlideCount, source.sourceSlideCount + 1);
  assert.deepEqual(source.nonTopologyChangedParts, []);
  assert.deepEqual(source.topologyChangedParts, ["[Content_Types].xml", "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"]);
  assert.equal(source.addedParts.some((name) => /^ppt\/slides\/slide[1-9][0-9]*\.xml$/u.test(name)), true);
  assert.equal(source.targetXmlMaskedEqual, true);
  assert.deepEqual(source.overlayRemovedParts, []);
  assert.deepEqual(source.unexpectedOverlayChanges, []);
  assert.equal(source.addedMediaParts.length, 1);
  assert.match(source.addedMediaParts[0], /^ppt\/media\/[^/]+\.png$/u);
  assert.equal(source.overlayAddedParts.includes(source.addedMediaParts[0]), true);
  assert.equal(source.overlayChangedExistingParts.includes(source.clonedSlidePart), true);
  assert.equal(source.overlayChangedExistingParts.includes(source.clonedSlidePart.replace(/\/([^/]+)$/u, "/_rels/$1.rels")), true);
  assert.match(source.sourceSha256, /^[a-f0-9]{64}$/u);
  assert.match(source.sourceSlideSha256, /^[a-f0-9]{64}$/u);
  assert.match(source.cloneOutputSha256, /^[a-f0-9]{64}$/u);
  assert.match(source.outputSha256, /^[a-f0-9]{64}$/u);
  assert.match(source.canonicalOpcSha256, /^[a-f0-9]{64}$/u);
  for (const hashes of [source.repeatOutputSha256s, source.repeatCanonicalOpcSha256s, source.repeatCloneOutputSha256s]) {
    assert.equal(hashes.length, 3);
    assert.equal(new Set(hashes).size, 1);
    assert.ok(hashes.every((hash) => /^[a-f0-9]{64}$/u.test(hash)));
  }
  assert.equal(source.repeatOutputSha256s[0], source.outputSha256);
  assert.equal(source.repeatCanonicalOpcSha256s[0], source.canonicalOpcSha256);
  assert.equal(source.repeatCloneOutputSha256s[0], source.cloneOutputSha256);
  assert.equal(source.repeatMutationFootprints.length, 3);
  assert.equal(new Set(source.repeatMutationFootprints).size, 1);
  assert.equal(source.target.capability.ready, true);
  assert.equal(source.target.capability.profile, "bounded-overlay");
  assert.equal(source.target.capability.embeddedImage, true);
  assert.equal(source.verifiedTarget.text.name, source.target.text.name);
  assert.equal(source.verifiedTarget.text.value, source.target.text.value);
  assert.equal(source.verifiedTarget.accent.name, source.target.accent.name);
  assert.equal(source.verifiedTarget.accent.geometry, "ellipse");
  assert.equal(source.verifiedTarget.image.name, source.target.image.name);
  assert.equal(source.verifiedTarget.image.alt, source.target.image.alt);
  assert.equal(source.verifiedTarget.image.sha256, source.target.image.sha256);
}
console.log("pptx source continuation ok");
