import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

const evidence = JSON.parse(await readFile(path.resolve("evals/pptx-lossless/source-continuation.v1.json"), "utf8"));
assert.equal(evidence.schema, "office-kit/pptx-source-continuation-evidence/v1");
assert.deepEqual(evidence.sources.map((source) => source.id), [
  "suanzhi-future-2026",
  "blue-gray-acid-template",
  "mckinsey-customer-loyalty",
]);
for (const source of evidence.sources) {
  assert.equal(source.sourceSlideUnchanged, true);
  assert.equal(source.outputSlideCount, source.sourceSlideCount + 1);
  assert.deepEqual(source.nonTopologyChangedParts, []);
  assert.deepEqual(source.topologyChangedParts, ["[Content_Types].xml", "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml"]);
  assert.equal(source.addedParts.some((name) => /^ppt\/slides\/slide[1-9][0-9]*\.xml$/u.test(name)), true);
  assert.match(source.sourceSha256, /^[a-f0-9]{64}$/u);
  assert.match(source.cloneOutputSha256, /^[a-f0-9]{64}$/u);
  assert.match(source.outputSha256, /^[a-f0-9]{64}$/u);
  if (source.kind === "text") {
    assert.equal(source.verifiedTarget.value, source.target.after);
    assert.notEqual(source.target.before, source.target.after);
  } else {
    assert.equal(source.verifiedTarget.marker, 'data-officekit="continuation"');
    assert.equal(source.addedParts.some((name) => /^ppt\/media\/image[1-9][0-9]*\.svg$/u.test(name)), true);
  }
}
console.log("pptx source continuation ok");
