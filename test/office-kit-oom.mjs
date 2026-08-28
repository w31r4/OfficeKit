import assert from "node:assert/strict";
import { createHash } from "node:crypto";

import {
  boundedInputBytes,
  codecLimits,
  invokeOfficeKit,
  invokeOfficeKitLazy,
  ownedInputBytes,
} from "../src/codecs/office-kit-runtime.mjs";
import { materializePresentationNativeGraphs } from "../src/codecs/office-kit-presentation-native.mjs";

const source = Uint8Array.from({ length: 4096 }, (_, index) => index & 0xff);
const strictLimits = codecLimits({ maxInputBytes: 2048 });

await assert.rejects(
  boundedInputBytes(source, strictLimits, "PPTX"),
  (error) => error?.name === "OfficeKitCodecError" && error?.code === "input_budget_exceeded",
  "input budgets must fail before codec invocation",
);
await assert.rejects(
  ownedInputBytes(source, strictLimits, "PPTX"),
  (error) => error?.name === "OfficeKitCodecError" && error?.code === "input_budget_exceeded",
  "owned source copies must happen only after the input budget passes",
);
await assert.rejects(
  invokeOfficeKit({ file: source, limits: strictLimits }),
  (error) => error?.name === "OfficeKitCodecError" && error?.code === "input_budget_exceeded",
  "the advanced codec boundary must enforce the same pre-encode budget",
);
await assert.rejects(
  invokeOfficeKitLazy(() => { throw new RangeError("Invalid array length"); }),
  (error) => error?.name === "OfficeKitCodecError" && error?.code === "js_memory_budget_exceeded",
  "catchable request-allocation failures must not be misreported as protocol failures",
);

const owned = await ownedInputBytes(source, { maxInputBytes: source.byteLength }, "PPTX");
assert.notEqual(owned.buffer, source.buffer, "accepted PPTX source bytes must remain mutation-safe");
const original = owned[0];
source[0] ^= 0xff;
assert.equal(owned[0], original, "the owned source snapshot must not alias caller memory");

const partPath = "ppt/charts/chart1.xml";
const partBytes = Uint8Array.from([1, 2, 3, 4]);
const partSha256 = createHash("sha256").update(partBytes).digest("hex");
const envelope = {
  opaqueOpc: {
    parts: [{ path: partPath, data: partBytes, sha256: partSha256, relationships: [] }],
    packageRelationships: [],
  },
  payload: {
    case: "presentation",
    value: {
      slides: [{ elements: [{ content: { case: "opaque", value: { preservedPartPaths: [partPath] } } }] }],
    },
  },
};
const nativeGraph = await materializePresentationNativeGraphs(envelope);
assert.deepEqual(
  [...nativeGraph({ preservedPartPaths: [partPath], relationshipReferences: [] }, "ppt/slides/slide1.xml").parts[0].bytes],
  [...partBytes],
  "sequential native-part materialization must preserve exact bytes",
);

console.log("OfficeKit JavaScript OOM guards ok");
