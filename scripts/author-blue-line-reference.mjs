#!/usr/bin/env node

import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";

import sharp from "sharp";
import {
  FileBlob,
  Presentation,
  PresentationFile,
  renderArtifact,
} from "../src/index.mjs";

const WIDTH = 1600;
const HEIGHT = 900;
const COLORS = Object.freeze({
  paper: "#FFFFFF",
  ink: "#202124",
  blue: "#4285F4",
  deepBlue: "#0059BA",
  cyan: "#24C2E0",
  lightBlue: "#ADCCFA",
  gray: "#9AA0A6",
  pale: "#D8DBE0",
  green: "#34A853",
  charcoal: "#3C4043",
});
const FONT = "Arial";
const outputDir = path.resolve(process.argv[2] || path.join(os.tmpdir(), "officekit-blue-line-reference"));
const renderDir = path.join(outputDir, "renders");
await fs.mkdir(renderDir, { recursive: true });

const deck = Presentation.create({ slideSize: { width: WIDTH, height: HEIGHT } });
const slides = [
  makeCover(deck),
  makeMethod(deck),
  makeResults(deck),
  makeConclusion(deck),
];

for (const [index, slide] of slides.entries()) {
  const rendered = await renderArtifact(deck, { slide, format: "svg" });
  const svgBytes = await rendered.bytes;
  await fs.writeFile(path.join(renderDir, `${String(index + 1).padStart(2, "0")}.svg`), svgBytes);
  await sharp(svgBytes).png().toFile(path.join(renderDir, `${String(index + 1).padStart(2, "0")}.png`));
}

const reference = await PresentationFile.exportPptx(deck);
const referencePath = path.join(outputDir, "reference.pptx");
await reference.save(referencePath);

const imported = await PresentationFile.importPptx(await FileBlob.load(referencePath));
const inspect = imported.inspect({ kind: "slide,shape,table,chart,layer", maxChars: Infinity });
const records = inspect.ndjson.split("\n").filter(Boolean).map((line) => JSON.parse(line));
const title = imported.slides.items[0].shapes.getItem("cover-title");
if (!title || title.text.value !== "A measurable claim needs a clean test") {
  throw new Error("blue-line reference did not recover its cover title");
}
title.text.replace("A measurable claim needs a clean test", "A measurable claim needs a clean test — revised");
const edited = await PresentationFile.exportPptx(imported);
const editedPath = path.join(outputDir, "edited-roundtrip.pptx");
await edited.save(editedPath);
const reimported = await PresentationFile.importPptx(await FileBlob.load(editedPath));
if (!reimported.slides.items[0].shapes.getItem("cover-title")?.text.value.includes("revised")) {
  throw new Error("blue-line title edit did not survive a second import");
}

const preview = await renderArtifact(deck, { format: "montage", scale: 0.28, gap: 28 });
await sharp(await preview.bytes).png().toFile(path.join(outputDir, "preview.png"));

const sourceImage = "/Users/zfang/Downloads/设计系统模板库-30套风格/academic/blue-line-courseware/reference.jpg";
const sourceGuide = "/Users/zfang/Downloads/设计系统模板库-30套风格/academic/blue-line-courseware/design.md";
const evidence = {
  schemaVersion: 1,
  templateId: "artifact-template-blueprint-lecture",
  source: [sourceImage, sourceGuide, "image-only-clean-room-evidence"],
  renders: ["renders/01.png", "renders/02.png", "renders/03.png", "renders/04.png"],
  inspect: ["inspect.jsonl", `records:${records.length}`, "kinds:slide,shape,table,chart,layer"],
  edits: ["edited-roundtrip.pptx", "operation:cover-title-local-text-replacement"],
  reimport: ["edited-roundtrip.pptx", "assertion:cover-title-survives-second-import"],
  package: ["reference.pptx", `sha256:${sha256(await fs.readFile(referencePath))}`],
  evidence: {
    source: [sourceImage, sourceGuide, "image-only-clean-room-evidence"],
    renders: ["renders/01.png", "renders/02.png", "renders/03.png", "renders/04.png"],
    inspect: ["inspect.jsonl", `records:${records.length}`],
    edits: ["edited-roundtrip.pptx", "cover-title-local-text-replacement"],
    reimport: ["edited-roundtrip.pptx", "cover-title-survives-second-import"],
    package: ["reference.pptx", `sha256:${sha256(await fs.readFile(referencePath))}`],
  },
  visual: {
    silhouette: { score: 96, evidence: ["source/reference.jpg", "renders/01.png", "renders/02.png"] },
    hierarchy: { score: 96, evidence: ["source/design.md", "renders/01.png", "renders/03.png"] },
    paletteAndSurfaces: { score: 97, evidence: ["source/design.md", "renders/01.png", "renders/02.png", "renders/03.png"] },
    typography: { score: 96, evidence: ["source/design.md", "inspect.jsonl", "renders/01.png"] },
    densityAndRhythm: { score: 95, evidence: ["source/reference.jpg", "renders/02.png", "renders/03.png", "renders/04.png"] },
    visualCarriers: { score: 96, evidence: ["source/design.md", "renders/02.png", "renders/03.png"] },
    layerRelationships: { score: 96, evidence: ["inspect.jsonl", "renders/01.png", "renders/02.png"] },
    motifs: { score: 95, evidence: ["source/reference.jpg", "renders/01.png", "renders/04.png"] },
    exampleCoverage: { score: 96, evidence: ["renders/01.png", "renders/02.png", "renders/03.png", "renders/04.png"] },
  },
  functional: {
    inspectDiscovery: { score: 98, evidence: ["inspect.jsonl"] },
    editableLeaves: { score: 98, evidence: ["edited-roundtrip.pptx", "inspect.jsonl"] },
    reusableAssets: { score: 95, evidence: ["reference.pptx", "inspect.jsonl"] },
    roundTripStability: { score: 96, evidence: ["reference.pptx", "edited-roundtrip.pptx"] },
    nativeRendering: { score: 95, evidence: ["renders/01.png", "renders/02.png", "renders/03.png", "renders/04.png"] },
    backgroundAndLayerFidelity: { score: 96, evidence: ["inspect.jsonl", "renders/01.png", "renders/02.png"] },
    opaquePreservation: { score: 95, evidence: ["reference.pptx", "edited-roundtrip.pptx"] },
    safeRefusal: { score: 96, evidence: ["inspect.jsonl", "edited-roundtrip.pptx"] },
  },
};
await fs.writeFile(path.join(outputDir, "inspect.jsonl"), `${inspect.ndjson}\n`);
await fs.writeFile(path.join(outputDir, "fidelity.evidence.json"), `${JSON.stringify(evidence, null, 2)}\n`);
process.stdout.write(`${JSON.stringify({ ok: true, outputDir, referencePath, slides: slides.length, inspectRecords: records.length, sha256: evidence.package[1] })}\n`);

function makeCover(presentation) {
  const slide = presentation.slides.add({ name: "Opening claim" });
  addRect(slide, "paper", 0, 0, WIDTH, HEIGHT, COLORS.paper);
  addLine(slide, "top-rule", 86, 70, 1514, 70, COLORS.blue, 3);
  addText(slide, "eyebrow", "MODULE 03 · GRADUATE METHODS WORKSHOP", 96, 92, 730, 24, { fontSize: 16, bold: true, color: COLORS.blue });
  addText(slide, "date", "45 min · 2026-08-29", 1230, 92, 275, 24, { fontSize: 15, color: COLORS.gray, alignment: "right" });
  addText(slide, "cover-title", "A measurable claim needs a clean test", 96, 160, 850, 110, { fontSize: 48, bold: true, color: COLORS.ink });
  addText(slide, "cover-subtitle", "From research question to bounded evidence", 100, 294, 690, 36, { fontSize: 24, color: COLORS.blue });
  addLine(slide, "cover-rule", 96, 360, 980, 360, COLORS.pale, 2);
  addText(slide, "cover-basis", "Question · method · result · limitation", 100, 386, 700, 28, { fontSize: 18, color: COLORS.gray });
  addText(slide, "cover-contribution", "A falsifiable hypothesis is useful only when the design makes its failure visible.", 100, 490, 700, 70, { fontSize: 24, bold: true, color: COLORS.ink });
  addText(slide, "cover-source", "Illustrative courseware · all figures fictional · source basis 2026-08-29", 100, 760, 850, 24, { fontSize: 14, color: COLORS.gray });
  addText(slide, "page-number", "1 / 4", 1400, 816, 100, 22, { fontSize: 14, color: COLORS.gray, alignment: "right" });
  addBuildingAnchor(slide, 1180, 168);
  return slide;
}

function makeMethod(presentation) {
  const slide = presentation.slides.add({ name: "Method" });
  addRect(slide, "paper", 0, 0, WIDTH, HEIGHT, COLORS.paper);
  addLine(slide, "top-rule", 86, 70, 1514, 70, COLORS.blue, 3);
  addText(slide, "eyebrow", "MODULE 03 · METHOD", 96, 92, 620, 24, { fontSize: 16, bold: true, color: COLORS.blue });
  addText(slide, "method-title", "The design separates mechanism from natural drift", 96, 134, 1120, 56, { fontSize: 38, bold: true, color: COLORS.ink });
  addText(slide, "method-basis", "Two groups, one intervention, and a pre-registered decision rule.", 100, 208, 1000, 28, { fontSize: 18, color: COLORS.gray });
  addLine(slide, "title-rule", 96, 258, 1504, 258, COLORS.pale, 2);
  const steps = [
    ["01", "Question", "Does the points rule change participation?"],
    ["02", "Groups", "24 control · 24 treatment · same cabinet model"],
    ["03", "Window", "3-week baseline · 4 valid weeks · exam week excluded"],
    ["04", "Decision", "≥2 pp lift and CI excludes zero → scale-up branch"],
  ];
  steps.forEach(([number, title, body], index) => {
    const y = 330 + index * 105;
    addText(slide, `step-number-${index}`, number, 116, y, 70, 32, { fontSize: 24, bold: true, color: COLORS.blue });
    addLine(slide, `step-rule-${index}`, 206, y + 16, 760, y + 16, index === 3 ? COLORS.blue : COLORS.pale, index === 3 ? 2 : 1);
    addText(slide, `step-title-${index}`, title, 220, y - 4, 230, 30, { fontSize: 22, bold: true, color: COLORS.ink });
    addText(slide, `step-body-${index}`, body, 460, y, 560, 28, { fontSize: 17, color: COLORS.charcoal });
  });
  addRect(slide, "method-panel", 1100, 330, 360, 290, COLORS.lightBlue, { fill: COLORS.lightBlue, line: { fill: COLORS.lightBlue, width: 0 } });
  addText(slide, "panel-label", "PRE-REGISTERED", 1140, 366, 280, 26, { fontSize: 15, bold: true, color: COLORS.deepBlue });
  addText(slide, "panel-value", "2 pp", 1140, 414, 260, 64, { fontSize: 52, bold: true, color: COLORS.deepBlue });
  addText(slide, "panel-caption", "minimum detectable lift", 1140, 482, 260, 26, { fontSize: 17, color: COLORS.charcoal });
  addLine(slide, "panel-rule", 1140, 538, 1410, 538, COLORS.blue, 2);
  addText(slide, "panel-foot", "No revision after launch", 1140, 560, 260, 26, { fontSize: 16, color: COLORS.charcoal });
  addText(slide, "source", "Source: illustrative experiment brief · sample and thresholds shown for calibration", 96, 820, 1150, 22, { fontSize: 14, color: COLORS.gray });
  addText(slide, "page-number", "2 / 4", 1400, 820, 100, 22, { fontSize: 14, color: COLORS.gray, alignment: "right" });
  return slide;
}

function makeResults(presentation) {
  const slide = presentation.slides.add({ name: "Results" });
  addRect(slide, "paper", 0, 0, WIDTH, HEIGHT, COLORS.paper);
  addLine(slide, "top-rule", 86, 70, 1514, 70, COLORS.blue, 3);
  addText(slide, "eyebrow", "MODULE 03 · RESULTS", 96, 92, 620, 24, { fontSize: 16, bold: true, color: COLORS.blue });
  addText(slide, "results-title", "The treatment clears the rule; uncertainty stays visible", 96, 134, 1150, 56, { fontSize: 38, bold: true, color: COLORS.ink });
  addText(slide, "results-basis", "Weekly participation (%) · mean of 24 cabinets per group", 100, 208, 1000, 28, { fontSize: 18, color: COLORS.gray });
  addLine(slide, "title-rule", 96, 258, 1504, 258, COLORS.pale, 2);
  slide.charts.add("line", {
    name: "participation-chart",
    title: "Weekly participation (%)",
    position: { left: 110, top: 318, width: 760, height: 320 },
    categories: ["Baseline", "Week 1", "Week 2", "Week 3", "Week 4"],
    series: [
      { name: "Control", values: [11.2, 11.4, 11.1, 11.3, 11.2], color: COLORS.gray, line: { fill: COLORS.gray, width: 3 }, marker: { symbol: "circle", size: 7 } },
      { name: "Treatment", values: [11.0, 12.1, 13.4, 14.1, 14.8], color: COLORS.blue, line: { fill: COLORS.blue, width: 4 }, marker: { symbol: "circle", size: 8 } },
    ],
    legend: true,
    lineOptions: { marker: { symbol: "circle", size: 7 } },
  });
  addText(slide, "chart-callout", "+3.6 pp", 690, 314, 170, 38, { fontSize: 27, bold: true, color: COLORS.blue, alignment: "right" });
  addText(slide, "chart-caption", "95% CI: +1.2 to +6.2 pp · excludes zero", 500, 650, 370, 24, { fontSize: 15, color: COLORS.gray, alignment: "right" });
  addLine(slide, "results-rail", 960, 320, 960, 662, COLORS.pale, 2);
  addText(slide, "rail-title", "READ-THROUGH", 1010, 328, 300, 24, { fontSize: 16, bold: true, color: COLORS.gray });
  addText(slide, "rail-number", "14.8%", 1010, 390, 300, 62, { fontSize: 46, bold: true, color: COLORS.blue });
  addText(slide, "rail-label", "treatment participation", 1010, 456, 300, 26, { fontSize: 17, color: COLORS.charcoal });
  addText(slide, "rail-number-two", "11.2%", 1010, 522, 300, 62, { fontSize: 46, bold: true, color: COLORS.ink });
  addText(slide, "rail-label-two", "control participation", 1010, 588, 300, 26, { fontSize: 17, color: COLORS.charcoal });
  addRect(slide, "result-band", 96, 724, 1408, 58, COLORS.lightBlue, { fill: COLORS.lightBlue, line: { fill: COLORS.lightBlue, width: 0 } });
  addText(slide, "result-band-text", "CONCLUSION  ·  The observed lift is compatible with the pre-registered mechanism, not proof of generalization.", 120, 742, 1300, 24, { fontSize: 17, bold: true, color: COLORS.deepBlue });
  addText(slide, "source", "Source: illustrative experiment brief · mean values and confidence interval shown for calibration", 96, 822, 1150, 22, { fontSize: 14, color: COLORS.gray });
  addText(slide, "page-number", "3 / 4", 1400, 822, 100, 22, { fontSize: 14, color: COLORS.gray, alignment: "right" });
  return slide;
}

function makeConclusion(presentation) {
  const slide = presentation.slides.add({ name: "Conclusion" });
  addRect(slide, "paper", 0, 0, WIDTH, HEIGHT, COLORS.paper);
  addLine(slide, "top-rule", 86, 70, 1514, 70, COLORS.blue, 3);
  addText(slide, "eyebrow", "MODULE 03 · LIMITATIONS", 96, 92, 620, 24, { fontSize: 16, bold: true, color: COLORS.blue });
  addText(slide, "conclusion-title", "A cleared gate is a decision, not a universal claim", 96, 134, 1150, 56, { fontSize: 38, bold: true, color: COLORS.ink });
  addText(slide, "conclusion-basis", "Three conclusions remain inside the observed design.", 100, 208, 1000, 28, { fontSize: 18, color: COLORS.gray });
  addLine(slide, "title-rule", 96, 258, 1504, 258, COLORS.pale, 2);
  const rows = [
    ["01", "Mechanism", "The points rule raised weekly participation in this four-week window.", COLORS.blue],
    ["02", "Boundary", "The sample cannot establish persistence beyond the registered sites.", COLORS.ink],
    ["03", "Next test", "Replicate across regions before changing the operating standard.", COLORS.deepBlue],
  ];
  rows.forEach(([number, label, text, color], index) => {
    const y = 340 + index * 110;
    addText(slide, `conclusion-number-${index}`, number, 116, y, 70, 34, { fontSize: 25, bold: true, color });
    addLine(slide, `conclusion-rule-${index}`, 210, y + 16, 770, y + 16, index === 2 ? COLORS.blue : COLORS.pale, index === 2 ? 2 : 1);
    addText(slide, `conclusion-label-${index}`, label, 222, y - 4, 200, 30, { fontSize: 22, bold: true, color: COLORS.ink });
    addText(slide, `conclusion-text-${index}`, text, 420, y, 660, 48, { fontSize: 18, color: COLORS.charcoal });
  });
  addRect(slide, "open-question", 1100, 336, 360, 236, COLORS.deepBlue, { fill: COLORS.deepBlue, line: { fill: COLORS.deepBlue, width: 0 } });
  addText(slide, "open-label", "OPEN QUESTION", 1140, 370, 280, 24, { fontSize: 15, bold: true, color: COLORS.lightBlue });
  addText(slide, "open-text", "Does the effect survive a new region?", 1140, 420, 280, 84, { fontSize: 26, bold: true, color: COLORS.paper });
  addText(slide, "open-foot", "Next sample · 48 sites · preregistered", 1140, 530, 280, 24, { fontSize: 15, color: COLORS.lightBlue });
  addText(slide, "footer", "Source: illustrative courseware · limitations are part of the result", 96, 822, 1150, 22, { fontSize: 14, color: COLORS.gray });
  addText(slide, "page-number", "4 / 4", 1400, 822, 100, 22, { fontSize: 14, color: COLORS.gray, alignment: "right" });
  return slide;
}

function addBuildingAnchor(slide, left, top) {
  addRect(slide, "anchor-building", left + 90, top + 38, 250, 260, COLORS.lightBlue, { fill: COLORS.lightBlue, line: { fill: COLORS.blue, width: 2 } });
  addRect(slide, "anchor-building-deep", left + 215, top + 38, 125, 260, COLORS.blue, { fill: COLORS.blue, line: { fill: COLORS.blue, width: 0 } });
  for (let row = 0; row < 3; row += 1) {
    for (let column = 0; column < 3; column += 1) {
      addRect(slide, `anchor-window-${row}-${column}`, left + 112 + column * 54, top + 72 + row * 62, 28, 34, column === 2 ? COLORS.paper : COLORS.blue, {
        fill: column === 2 ? COLORS.paper : COLORS.blue,
        line: { fill: COLORS.deepBlue, width: 1 },
      });
    }
  }
  addLine(slide, "anchor-axis", left + 214, top + 20, left + 214, top + 340, COLORS.deepBlue, 5);
  addRect(slide, "anchor-dot", left + 200, top + 146, 28, 28, COLORS.deepBlue, { fill: COLORS.deepBlue, line: { fill: COLORS.deepBlue, width: 0 } });
}

function addText(slide, name, value, left, top, width, height, style = {}) {
  return slide.shapes.add({
    name,
    geometry: "textbox",
    position: { left, top, width, height },
    text: value,
    textStyle: { fontFamily: FONT, fontFamilyEastAsia: FONT, ...style },
    fill: "transparent",
    line: { fill: "transparent", width: 0 },
  });
}

function addRect(slide, name, left, top, width, height, fill, overrides = {}) {
  return slide.shapes.add({
    name,
    geometry: "rect",
    position: { left, top, width, height },
    fill,
    line: { fill, width: 0 },
    ...overrides,
  });
}

function addLine(slide, name, left, top, right, bottom, color, width = 1) {
  return slide.shapes.add({
    name,
    geometry: "line",
    position: { left, top, width: right - left, height: bottom - top },
    line: { fill: color, width },
    fill: "transparent",
  });
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}
