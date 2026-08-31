import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";

import { reviewArtifact } from "../src/review/index.mjs";

const bytes = Uint8Array.from([1, 2, 3, 4]);
const outputSha256 = createHash("sha256").update(bytes).digest("hex");
const program = {
  schema: "office-kit/ppj/v1",
  design: { canvas: { width: 960, height: 540, unit: "pt" } },
  pages: [{
    id: "page-1",
    name: "Review",
    elements: [{ id: "title", type: "text", frame: { x: 48, y: 48, width: 864, height: 80 }, text: "PPJ review" }],
  }],
};
const ppjReceipt = {
  programJson: Buffer.from(JSON.stringify(program)),
  programSha256: createHash("sha256").update(JSON.stringify(program)).digest("hex"),
  nodeMapJson: Buffer.from("{}"),
  outputSha256,
  sourceSha256: "",
  sourceBound: false,
  restoredEmbeddedProgram: true,
  expandedElementCount: 1,
  diagnostics: [],
};

const report = await reviewArtifact(bytes, {
  format: "pptx",
  outputPath: "tmp/review.ppj.pptx",
  ppjReceipt,
  contentView: "none",
  visualReview: "complete",
});
assert.equal(report.verdict, "passed");
assert.equal(report.semantic.recordCounts.page, 1);
assert.equal(report.semantic.recordCounts["element:text"], 1);
assert.equal(report.structural.summary.programSha256, ppjReceipt.programSha256);
assert.equal(report.layout.scope, "PPJ frames and NativeAOT projection");

const source = await readFile(new URL("../src/review/index.mjs", import.meta.url), "utf8");
assert.doesNotMatch(source, /presentation\/index\.mjs|PresentationFile|instanceof Presentation/u);

console.log("PPJ-native review smoke ok");
