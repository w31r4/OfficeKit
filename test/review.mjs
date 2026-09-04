import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";

import { reviewPpjArtifact } from "../src/ppj/review.mjs";
import { reviewArtifact } from "../src/review/index.mjs";

const bytes = Uint8Array.from([1, 2, 3, 4]);
const sha256 = (value) => createHash("sha256").update(value).digest("hex");
const outputSha256 = createHash("sha256").update(bytes).digest("hex");
const program = {
  schema: "office-kit/ppj/v1",
  design: { canvas: { width: 960, height: 540, unit: "pt" } },
  pages: [{
    id: "page-1",
    name: "Review",
    readingOrder: ["title"],
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
assert.equal(report.layout.scope, "PPJ visual bounds, z-order, and overlap (read-only review)");
assert.equal(report.accessibility.scope, "PPJ accessibility metadata and reading order (machine review)");
assert.equal(report.accessibility.pages[0].source, "explicit");

const rotatedProgram = {
  schema: "office-kit/ppj/v1",
  design: {
    canvas: { width: 960, height: 540, unit: "pt" },
    grammar: {
      tokens: { accent: { kind: "color", value: "#336699" } },
      stylePrecedence: [{ target: "fill.color", sources: ["inline", "styleRef", "default"] }],
      predicates: [{ id: "visual-elements", field: "element.type", op: "in", value: ["shape", "text", "line"] }],
    },
  },
  pages: [{
    id: "page-1",
    name: "Visual bounds",
    elements: [
      { id: "rotated", type: "shape", frame: { x: 100, y: 100, width: 100, height: 100, rotation: 45 }, style: { fill: { color: { token: "accent" } } } },
      { id: "neighbor", type: "shape", frame: { x: 210, y: 150, width: 50, height: 50 } },
      { id: "shadow-edge", type: "shape", frame: { x: 900, y: 80, width: 50, height: 50 }, style: { shadow: { color: "#000000", blur: 10, distance: 30, angle: 0 } } },
      { id: "arrow-line", type: "line", frame: { x: 100, y: 0, width: 100, height: 2 }, viewBox: [100, 2], points: "0,1 100,1", curve: "sharp", stroke: { color: "#000000", width: 2 }, endArrow: "triangle" },
      { id: "arrow-neighbor", type: "shape", frame: { x: 190, y: -3, width: 30, height: 20 } },
      { id: "text-overflow", type: "text", frame: { x: 48, y: 300, width: 36, height: 14 }, text: "A long label that cannot fit", textStyle: { margins: { left: 2, right: 2, top: 1, bottom: 1 }, autoFit: "none" } },
    ],
  }],
};
const rotatedJson = JSON.stringify(rotatedProgram);
const rotatedReceipt = {
  ...ppjReceipt,
  programJson: Buffer.from(rotatedJson),
  programSha256: createHash("sha256").update(rotatedJson).digest("hex"),
  expandedElementCount: 6,
};
const visualBoundsReport = await reviewPpjArtifact(bytes, { ppjReceipt: rotatedReceipt });
const rotatedOverlap = visualBoundsReport.layout.issues.find((entry) => entry.type === "elementOverlap" && entry.ids?.includes("rotated"));
assert.equal(rotatedOverlap?.detection, "rotated-visual-bounds");
assert.ok(rotatedOverlap?.visualBounds?.[0]?.[2] > 100);
const shadowOverflow = visualBoundsReport.layout.issues.find((entry) => entry.type === "frameOutsideCanvas" && entry.id === "shadow-edge");
assert.equal(shadowOverflow?.visualBounds?.shadow, true);
assert.ok(shadowOverflow?.bbox?.[2] > 50);
const arrowOverflow = visualBoundsReport.layout.issues.find((entry) => entry.type === "frameOutsideCanvas" && entry.id === "arrow-line");
assert.equal(arrowOverflow?.visualBounds?.arrowheads, true);
assert.deepEqual(arrowOverflow?.visualBounds?.arrowKinds, ["triangle"]);
assert.ok(arrowOverflow?.bbox?.[1] < 0);
assert.match(arrowOverflow?.recommendation || "", /arrowhead visual bounds/u);
const arrowOverlap = visualBoundsReport.layout.issues.find((entry) => entry.type === "elementOverlap" && entry.ids?.includes("arrow-line"));
assert.equal(arrowOverlap?.detection, "arrow-visual-bounds");
assert.match(arrowOverlap?.recommendation || "", /arrow endpoint/u);
const textOverflow = visualBoundsReport.layout.issues.find((entry) => entry.type === "textOverflowEstimated" && entry.id === "text-overflow");
assert.equal(textOverflow?.severity, "warning");
assert.equal(textOverflow?.measurement?.method, "deterministic-character-metric");
assert.equal(visualBoundsReport.design.grammar.tokens.accent.value, "#336699");
assert.equal(visualBoundsReport.design.grammar.resolutions[0].entries[0].source, "inline");
assert.equal(visualBoundsReport.design.grammar.resolutions[0].entries[0].value, "#336699");
assert.equal(visualBoundsReport.design.grammar.resolutions[0].entries[0].token, "accent");
assert.equal(visualBoundsReport.design.grammar.predicates[0].checks.every((entry) => entry.pass), true);

const accessibilityProgram = {
  schema: "office-kit/ppj/v1",
  design: { canvas: { width: 960, height: 540, unit: "pt" } },
  pages: [{
    id: "page-a11y",
    readingOrder: ["figure", "figure"],
    elements: [
      { id: "figure", type: "image", frame: { x: 10, y: 10, width: 100, height: 100 }, asset: "figure" },
      { id: "caption", type: "text", frame: { x: 10, y: 120, width: 100, height: 20 }, text: "Caption" },
    ],
  }],
};
const accessibilityJson = JSON.stringify(accessibilityProgram);
const accessibilityReport = await reviewPpjArtifact(bytes, {
  ppjReceipt: { ...ppjReceipt, programJson: Buffer.from(accessibilityJson), programSha256: sha256(accessibilityJson), expandedElementCount: 2 },
});
assert.equal(accessibilityReport.accessibility.status, "failed");
assert.ok(accessibilityReport.accessibility.issues.some((entry) => entry.type === "invalidReadingOrder" && entry.severity === "error"));
assert.ok(accessibilityReport.accessibility.issues.some((entry) => entry.type === "missingAlternativeText" && entry.id === "figure"));

const source = await readFile(new URL("../src/review/index.mjs", import.meta.url), "utf8");
assert.doesNotMatch(source, /presentation\/index\.mjs|PresentationFile|instanceof Presentation/u);

console.log("PPJ-native review smoke ok");
