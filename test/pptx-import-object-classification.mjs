import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { FileBlob, Presentation, PresentationFile } from "../src/index.mjs";
import { classifyImportedPresentationObjects } from "../src/presentation/import-object-classification.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const fixture = path.resolve(import.meta.dirname, "../evals/assets/presentations/strategy-review.pptx");
const sourceBytes = await readFile(fixture);

assert.throws(
  () => Presentation.create().inspect({ kind: "importObject" }),
  (error) => error?.code === "presentation_import_object_source_required",
);

const imported = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME }));
const first = records(imported.inspect({ includeImportObjects: true, maxChars: Infinity }).ndjson)
  .filter((record) => record.kind === "importObject");
const second = records(imported.inspect({ kind: "importObject", maxChars: Infinity }).ndjson);
assert.deepEqual(first, second, "imported-object classification must be deterministic across inspect modes");
assert.equal(first.length, 10);
assert.equal(first.every((record) => record.topLevel === true), true);
assert.equal(first.every((record) => record.classification === "typed-editable"), true);
assert.equal(first.every((record) => record.typedOperations.length > 0), true);
assert.equal(first.every((record) => /^[0-9a-f]{64}$/u.test(record.sourceRevisionSha256)), true);
assert.equal(new Set(first.map((record) => record.sourceRevisionSha256)).size, 1);
assert.equal(new Set(first.map((record) => record.id)).size, first.length);
assert.equal(new Set(first.map((record) => record.targetId)).size, first.length);
assert.equal(new Set(first.map((record) => `${record.sourceLocator.slideId}:${record.sourceLocator.shapeTreeIndex}`)).size, first.length);
assert.equal(first.every((record) => /^[0-9a-f]{64}$/u.test(record.sourceLocator.expectedElementSha256)), true);
assert.equal(first.every((record) => /^[0-9a-f]{64}$/u.test(record.sourceLocator.expectedSemanticSha256)), true);
assert.equal(first.some((record) => record.nativeLeafKinds.includes("diagramText")), true);
assert.equal(first.some((record) => record.reuse.length > 0), true);
assert.equal(/rawXml|partPath|relationshipId|sourcePackage|<p:/u.test(JSON.stringify(first)), false);

const target = first[0];
assert.deepEqual(
  records(imported.inspect({ kind: "importObject", target: target.targetId, maxChars: Infinity }).ndjson),
  [target],
);
const mutatedRecord = structuredClone(target);
mutatedRecord.classification = "opaque-preserved";
assert.equal(records(imported.inspect({ kind: "importObject", target: target.targetId, maxChars: Infinity }).ndjson)[0].classification, "typed-editable");
const noOp = await PresentationFile.exportPptx(imported);
assert.deepEqual([...noOp.bytes], [...sourceBytes], "classification inspection must preserve exact no-op bytes");

const fakeState = classificationFixtureState();
const classified = classifyImportedPresentationObjects(fakeState, {
  nativeLeafRecords: [{ kind: "nativeLeaf", leafId: "nl_native", targetId: "slide-1/native", leafKind: "leftEmu" }],
  componentRecords: [{
    candidateId: "pc_reusable",
    occurrences: [{ targetId: "slide-1/reusable", reuseCapability: { supported: true } }],
  }],
});
assert.deepEqual(classified.map((record) => record.classification), [
  "typed-editable",
  "native-leaf-editable",
  "source-derived-reusable",
  "opaque-preserved",
]);
assert.deepEqual(classified[0].typedOperations, ["semantic-model"]);
assert.deepEqual(classified[1].nativeLeafKinds, ["leftEmu"]);
assert.deepEqual(classified[2].reuse, [{ candidateId: "pc_reusable", targetId: "slide-1/reusable" }]);
assert.match(classified[3].reason, /outside the controlled profile/u);

const duplicate = classificationFixtureState();
duplicate.slides[0].entries[1].wire.source.shapeTreeIndex = 0;
assert.throws(
  () => classifyImportedPresentationObjects(duplicate),
  (error) => error?.code === "presentation_import_object_binding_invalid",
);

console.log("PPTX imported-object classification smoke ok");

function records(ndjson) {
  return ndjson.trim().split("\n").filter(Boolean).map((line) => JSON.parse(line));
}

function classificationFixtureState() {
  const revision = "a".repeat(64);
  const makeEntry = (id, index, contentCase, model = {}, source = {}) => ({
    wire: {
      id,
      content: { case: contentCase, value: contentCase === "opaque" ? { nativeKind: model.nativeKind || "opaque" } : {} },
      source: {
        shapeTreeIndex: index,
        elementSha256: String(index + 1).repeat(64).slice(0, 64),
        semanticSha256: String(index + 5).repeat(64).slice(0, 64),
        editable: false,
        textEditable: false,
        accessibilityEditable: false,
        ...source,
      },
    },
    model: { id, name: id, ...model },
  });
  return {
    source: { packageSha256: revision },
    opaqueOpc: { sourcePackage: { sha256: revision } },
    slides: [{
      wire: { id: "slide-1" },
      slide: { index: 0 },
      entries: [
        makeEntry("slide-1/typed", 0, "shape", {}, { editable: true }),
        makeEntry("slide-1/native", 1, "opaque", { nativeKind: "picture" }),
        makeEntry("slide-1/reusable", 2, "shape"),
        makeEntry("slide-1/opaque", 3, "opaque", {
          nativeKind: "group",
          deletionCapability: { supported: false, blockedReason: "group topology is outside the controlled profile" },
        }),
      ],
    }],
  };
}
