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

const formalReviewProgram = {
  schema: "office-kit/ppj/v1",
  design: {
    theme: { colors: [], textStyle: { bold: true, fontFamilyEastAsia: "Noto Sans CJK SC", language: { token: "language" } } },
    styles: {
      text: [{ id: "formal-named", style: { defaultText: { size: 16 } } }],
    },
    masters: [
      { id: "formal-master", style: { defaultText: { size: 12 } } },
      { id: "empty-master" },
    ],
    layouts: [
      { id: "formal-layout", master: "formal-master", style: { defaultText: { size: 14 } } },
      { id: "master-layout", master: "formal-master" },
      { id: "default-layout", master: "empty-master" },
    ],
    grammar: {
      tokens: {
        size: { kind: "size", value: 9 },
        bold: { kind: "boolean", value: false },
        fontFamilyEastAsia: { kind: "string", value: "Source Han Sans SC" },
        language: { kind: "string", value: "ar-SA" },
      },
      stylePrecedence: [
        {
          target: "text.size",
          sources: ["run", "paragraph", "element", "styleRef", "layout", "master", "theme", "default"],
        },
        {
          target: "text.bold",
          sources: ["run", "paragraph", "element", "styleRef", "layout", "master", "theme", "default"],
        },
        {
          target: "text.fontFamilyEastAsia",
          sources: ["run", "paragraph", "element", "styleRef", "layout", "master", "theme", "default"],
        },
        {
          target: "text.language",
          sources: ["run", "paragraph", "element", "styleRef", "layout", "master", "theme", "default"],
        },
      ],
    },
  },
  pages: [
    {
      id: "formal-page",
      layout: "formal-layout",
      elements: [
        {
          id: "run-wins",
          type: "text",
          styleRef: "formal-named",
          textStyle: { size: 18 },
          text: { paragraphs: [{ style: { defaultText: { size: 20 } }, runs: [{ text: "run", style: { size: 22 } }] }] },
        },
        {
          id: "paragraph-wins",
          type: "text",
          styleRef: "formal-named",
          textStyle: { size: 18 },
          text: { paragraphs: [{ style: { defaultText: { size: 20 } }, runs: [{ text: "paragraph" }] }] },
        },
        {
          id: "element-wins",
          type: "text",
          styleRef: "formal-named",
          textStyle: { size: 18 },
          text: { paragraphs: [{ runs: [{ text: "element" }] }] },
        },
        {
          id: "named-wins",
          type: "text",
          styleRef: "formal-named",
          text: { paragraphs: [{ runs: [{ text: "named" }] }] },
        },
        {
          id: "layout-wins",
          type: "text",
          text: { paragraphs: [{ runs: [{ text: "layout" }] }] },
        },
      ],
    },
    {
      id: "master-page",
      layout: "master-layout",
      elements: [{ id: "master-wins", type: "text", text: { paragraphs: [{ runs: [{ text: "master" }] }] } }],
    },
    {
      id: "default-page",
      layout: "default-layout",
      elements: [{ id: "default-wins", type: "text", text: { paragraphs: [{ runs: [{ text: "default" }] }] } }],
    },
  ],
};
const formalReviewJson = JSON.stringify(formalReviewProgram);
const formalReview = await reviewPpjArtifact(bytes, {
  ppjReceipt: {
    ...ppjReceipt,
    programJson: Buffer.from(formalReviewJson),
    programSha256: sha256(formalReviewJson),
    expandedElementCount: 7,
  },
});
const formalResolution = formalReview.design.grammar.resolutions.find((entry) => entry.target === "text.size");
assert.deepEqual(formalResolution.entries.map(({ id, source, value, paragraphIndex, runIndex }) => ({
  id,
  source,
  value,
  paragraphIndex,
  runIndex,
})), [
  { id: "run-wins", source: "run", value: 22, paragraphIndex: 0, runIndex: 0 },
  { id: "paragraph-wins", source: "paragraph", value: 20, paragraphIndex: 0, runIndex: 0 },
  { id: "element-wins", source: "element", value: 18, paragraphIndex: 0, runIndex: 0 },
  { id: "named-wins", source: "styleRef", value: 16, paragraphIndex: 0, runIndex: 0 },
  { id: "layout-wins", source: "layout", value: 14, paragraphIndex: 0, runIndex: 0 },
  { id: "master-wins", source: "master", value: 12, paragraphIndex: 0, runIndex: 0 },
  { id: "default-wins", source: "default", value: 9, paragraphIndex: 0, runIndex: 0 },
]);
const formalBoldResolution = formalReview.design.grammar.resolutions.find((entry) => entry.target === "text.bold");
assert.equal(formalBoldResolution.entries.length, 7);
assert.ok(formalBoldResolution.entries.every(({ source, value }) => source === "theme" && value === true));
const formalEastAsiaResolution = formalReview.design.grammar.resolutions.find((entry) => entry.target === "text.fontFamilyEastAsia");
assert.equal(formalEastAsiaResolution.entries.length, 7);
assert.ok(formalEastAsiaResolution.entries.every(({ source, value }) => source === "theme" && value === "Noto Sans CJK SC"));
const formalLanguageResolution = formalReview.design.grammar.resolutions.find((entry) => entry.target === "text.language");
assert.equal(formalLanguageResolution.entries.length, 7);
assert.ok(formalLanguageResolution.entries.every(({ source, value }) => source === "theme" && value === "ar-SA"));
const formalDefaultProgram = structuredClone(formalReviewProgram);
delete formalDefaultProgram.design.theme.textStyle.fontFamilyEastAsia;
const formalDefaultJson = JSON.stringify(formalDefaultProgram);
const formalDefaultReview = await reviewPpjArtifact(bytes, {
  ppjReceipt: {
    ...ppjReceipt,
    programJson: Buffer.from(formalDefaultJson),
    programSha256: sha256(formalDefaultJson),
    expandedElementCount: 7,
  },
});
const formalDefaultEastAsiaResolution = formalDefaultReview.design.grammar.resolutions.find((entry) => entry.target === "text.fontFamilyEastAsia");
assert.equal(formalDefaultEastAsiaResolution.entries.length, 7);
assert.ok(formalDefaultEastAsiaResolution.entries.every(({ source, value }) => source === "default" && value === "Source Han Sans SC"));

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
