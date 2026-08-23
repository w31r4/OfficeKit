import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { buildImportObjectClassificationEvidence } from "../scripts/pptx-import-object-classification.mjs";

const evidencePath = path.resolve(import.meta.dirname, "../evals/pptx-lossless/import-object-classification.v1.json");
const evidence = JSON.parse(await readFile(evidencePath, "utf8"));

assert.equal(evidence.schema, "office-kit/pptx-import-object-classification-evidence/v1");
assert.deepEqual(evidence.oracle, {
  rawSourceIndependent: true,
  directShapeTreeChildren: true,
  runtimeSelfReportedCountInsufficient: true,
  exactNoOpBytes: true,
});
assert.deepEqual(evidence.totals, {
  sources: 3,
  slides: 48,
  visibleTopLevelObjects: 926,
  classifiedTopLevelObjects: 926,
});
assert.equal(evidence.sources.length, 3);
assert.deepEqual(evidence.sources.map((source) => source.id), [
  "suanzhi-future-2026",
  "blue-gray-acid-template",
  "mckinsey-customer-loyalty",
]);

for (const source of evidence.sources) {
  assert.equal(source.complete, true);
  assert.equal(source.noOpByteIdentical, true);
  assert.equal(source.visibleTopLevelObjects, source.classifiedTopLevelObjects);
  assert.equal(source.slideCount, source.slides.length);
  assert.equal(source.slides.reduce((sum, slide) => sum + slide.visibleTopLevelObjects, 0), source.visibleTopLevelObjects);
  assert.equal(source.slides.every((slide) => slide.visibleTopLevelObjects === slide.classifiedTopLevelObjects), true);
  assert.equal(Object.values(source.classifications).reduce((sum, count) => sum + count, 0), source.classifiedTopLevelObjects);
  assert.match(source.sha256, /^[0-9a-f]{64}$/u);
  assert.match(source.slidesEvidenceSha256, /^[0-9a-f]{64}$/u);
}

const suanzhi = evidence.sources[0];
assert.equal(suanzhi.slideCount, 21);
assert.equal(suanzhi.visibleTopLevelObjects, 538);
assert.equal(suanzhi.rawRootKinds.graphicFrame, 22);
assert.equal(suanzhi.objectKinds.oleObject, 18);
assert.equal(suanzhi.classifications["opaque-preserved"], 19);

const blueGray = evidence.sources[1];
assert.equal(blueGray.slideCount, 19);
assert.equal(blueGray.visibleTopLevelObjects, 380);
assert.equal(blueGray.rawRootKinds.grpSp, 99);
assert.equal(blueGray.classifications["native-leaf-editable"], 6);
assert.equal(blueGray.classifications["opaque-preserved"], 4);

const mckinsey = evidence.sources[2];
assert.equal(mckinsey.slideCount, 8);
assert.equal(mckinsey.visibleTopLevelObjects, 8);
assert.deepEqual(mckinsey.rawRootKinds, { pic: 8 });
assert.deepEqual(mckinsey.objectKinds, { image: 8 });
assert.equal(mckinsey.typedOperations["svg-text"], 8);

const assetsDir = process.env.OFFICEKIT_PPTX_BENCHMARK_ASSETS;
if (assetsDir) {
  const rebuilt = await buildImportObjectClassificationEvidence({ assetsDir });
  assert.deepEqual(rebuilt, evidence, "real PPTX classification evidence drifted from the immutable sources");
}

console.log(`PPTX imported-object evidence smoke ok${assetsDir ? " (real sources rebuilt)" : " (committed evidence)"}`);
