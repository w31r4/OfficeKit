import assert from "node:assert/strict";
import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import {
  PresentationProgramRequestSchema, PresentationProgramResultSchema,
  PresentationPreviewSceneOrigin, PresentationPreviewAttribution,
} from "../src/generated/office_kit/artifact/v1/office_artifact_pb.js";

const roundTrip = (schema, message) => fromBinary(schema, toBinary(schema, message));
assert.equal(create(PresentationProgramRequestSchema).includePreviewScene, false);
assert.equal(create(PresentationProgramResultSchema).previewScene, undefined);
assert.deepEqual(toBinary(PresentationProgramRequestSchema, create(PresentationProgramRequestSchema)), new Uint8Array());
assert.equal(roundTrip(PresentationProgramRequestSchema,
  create(PresentationProgramRequestSchema, { includePreviewScene: true })).includePreviewScene, true);

const receipt = create(PresentationProgramResultSchema, {
  programJson: new TextEncoder().encode('{ "canonical": true }'),
  programSha256: "a".repeat(64), outputSha256: "b".repeat(64),
  previewScene: {
    version: 1, origin: PresentationPreviewSceneOrigin.AUTHORED_LOWERING,
    programSha256: "a".repeat(64), candidateSha256: "b".repeat(64),
    presentation: { id: "deck", slides: [{ id: "page", hidden: false, elements: [
      { id: "explicit", hidden: false, locked: false, content: { case: "chart", value: {
        yAxis: { minimum: 0, maximum: 10, reverse: false },
      } } },
      { id: "absent" },
    ] }] },
    bindings: [{ pageId: "page", semanticId: "explicit", nativeId: "explicit",
      programPath: "$.pages[0].elements[0]", scenePath: "$.presentation.slides[0].elements[0]",
      attribution: PresentationPreviewAttribution.DIRECT, zOrder: 0 }],
    assets: [{ nativeId: "native-image", contentType: "image/png", sha256: "c".repeat(64) }],
  },
});
const restored = roundTrip(PresentationProgramResultSchema, receipt);
assert.deepEqual(restored, receipt);
assert.deepEqual(restored.programJson, receipt.programJson);
const [explicit, absent] = restored.previewScene.presentation.slides[0].elements;
assert.equal(explicit.hidden, false);
assert.equal(explicit.locked, false);
assert.equal(absent.hidden, undefined);
assert.equal(absent.locked, undefined);
assert.equal(explicit.content.value.yAxis.minimum, 0);
assert.equal(explicit.content.value.yAxis.reverse, false);
assert.equal(explicit.content.value.yAxis.minorUnit, undefined);
assert.equal(restored.previewScene.bindings[0].zOrder, 0);
// Asset references are identities, not a second payload transport.
assert.equal(Object.hasOwn(restored.previewScene.assets[0], "data"), false);
assert.equal(Object.hasOwn(restored.previewScene, "opaqueOpc"), false);
console.log("ppj preview scene wire ok: opt-in, existing receipts, native optional presence and identity references");
