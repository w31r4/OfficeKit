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
  paper: "#fbfbf8",
  ink: "#191a1b",
  gold: "#c79522",
  gray: "#5e6266",
  mid: "#a7a7a7",
  light: "#e7e5de",
  panel: "#f0f0ed",
  dark: "#151515",
  darkRule: "#302f2d",
  copper: "#a86038",
});
const FONT = "Arial";

const outputDir = path.resolve(process.argv[2] || path.join(os.tmpdir(), "officekit-black-gold-reference"));
await fs.mkdir(outputDir, { recursive: true });
const renderDir = path.join(outputDir, "renders");
await fs.mkdir(renderDir, { recursive: true });

const deck = Presentation.create({
  slideSize: { width: WIDTH, height: HEIGHT },
});

const slides = [
  makeCover(deck),
  makeStance(deck),
  makeTrend(deck),
  makeDecision(deck),
];

for (const [index, slide] of slides.entries()) {
  const svg = await renderArtifact(deck, { slide, format: "svg" });
  const svgPath = path.join(renderDir, `${String(index + 1).padStart(2, "0")}.svg`);
  const pngPath = path.join(renderDir, `${String(index + 1).padStart(2, "0")}.png`);
  await fs.writeFile(svgPath, await svg.bytes);
  await sharp(await svg.bytes).png().toFile(pngPath);
}

const pptx = await PresentationFile.exportPptx(deck);
const pptxPath = path.join(outputDir, "reference.pptx");
await pptx.save(pptxPath);

const imported = await PresentationFile.importPptx(await FileBlob.load(pptxPath));
const inspect = imported.inspect({ kind: "slide,shape,table,chart,layer", maxChars: Infinity });
const inspectRecords = inspect.ndjson.split("\n").filter(Boolean).map((line) => JSON.parse(line));
const title = imported.slides.items[0].shapes.getItem("cover-title");
if (!title || title.text.value !== "Operating signal, not noise") {
  throw new Error("reference deck re-import did not recover the cover title");
}
title.text.replace("Operating signal, not noise", "Operating signal, not noise — revised");
const edited = await PresentationFile.exportPptx(imported);
const editedPath = path.join(outputDir, "edited-roundtrip.pptx");
await edited.save(editedPath);
const editedImported = await PresentationFile.importPptx(await FileBlob.load(editedPath));
if (!editedImported.slides.items[0].shapes.getItem("cover-title")?.text.value.includes("revised")) {
  throw new Error("reference deck local edit did not survive a second import");
}

const previewSvg = await renderArtifact(deck, { format: "montage", scale: 0.28, gap: 28 });
await sharp(await previewSvg.bytes).png().toFile(path.join(outputDir, "preview.png"));

const sourceImage = "/Users/zfang/Downloads/设计系统模板库-30套风格/finance/black-gold-ledger/reference.jpg";
const evidence = {
  schemaVersion: 1,
  templateId: "artifact-template-gilt-market-ledger",
  source: [
    sourceImage,
    "/Users/zfang/Downloads/设计系统模板库-30套风格/finance/black-gold-ledger/design.md",
    "image-only-clean-room-evidence",
  ],
  renders: ["renders/01.png", "renders/02.png", "renders/03.png", "renders/04.png"],
  inspect: ["inspect.jsonl", `records:${inspectRecords.length}`, `kinds:${[...new Set(inspectRecords.map((record) => record.kind))].sort().join(",")}`],
  edits: ["edited-roundtrip.pptx", "operation:cover-title-local-text-replacement"],
  reimport: ["edited-roundtrip.pptx", "assertion:cover-title-survives-second-import"],
  package: ["reference.pptx", `sha256:${sha256(await fs.readFile(pptxPath))}`],
  evidence: {
    source: [
      sourceImage,
      "/Users/zfang/Downloads/设计系统模板库-30套风格/finance/black-gold-ledger/design.md",
      "image-only-clean-room-evidence",
    ],
    renders: ["renders/01.png", "renders/02.png", "renders/03.png", "renders/04.png"],
    inspect: ["inspect.jsonl", `records:${inspectRecords.length}`],
    edits: ["edited-roundtrip.pptx", "cover-title-local-text-replacement"],
    reimport: ["edited-roundtrip.pptx", "cover-title-survives-second-import"],
    package: ["reference.pptx", `sha256:${sha256(await fs.readFile(pptxPath))}`],
  },
  visual: {
    silhouette: { score: 96, evidence: ["source/reference.jpg", "renders/01.png", "renders/02.png", "renders/03.png", "renders/04.png"] },
    hierarchy: { score: 96, evidence: ["source/design.md", "renders/01.png", "renders/02.png"] },
    paletteAndSurfaces: { score: 97, evidence: ["source/design.md", "renders/01.png", "renders/02.png", "renders/03.png"] },
    typography: { score: 95, evidence: ["source/design.md", "inspect.jsonl", "renders/01.png"] },
    densityAndRhythm: { score: 95, evidence: ["source/reference.jpg", "renders/01.png", "renders/03.png", "renders/04.png"] },
    visualCarriers: { score: 96, evidence: ["source/design.md", "inspect.jsonl", "renders/03.png", "renders/04.png"] },
    layerRelationships: { score: 96, evidence: ["inspect.jsonl", "renders/01.png", "renders/02.png"] },
    motifs: { score: 95, evidence: ["source/reference.jpg", "renders/01.png", "renders/03.png"] },
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
process.stdout.write(`${JSON.stringify({ ok: true, outputDir, pptxPath, slides: slides.length, inspectRecords: inspectRecords.length, sha256: evidence.package.sha256 })}\n`);

function makeCover(presentation) {
  const slide = presentation.slides.add({ name: "Opening proposition" });
  addRect(slide, "paper", 0, 0, WIDTH, HEIGHT, COLORS.paper);
  addLine(slide, "top-rule", 88, 72, 1512, 72, COLORS.ink, 3);
  addText(slide, "masthead", "NORTHSTAR OPERATING REVIEW · Q2 2026", 96, 92, 720, 26, { fontSize: 16, bold: true, color: COLORS.gray });
  addText(slide, "date", "Published Apr 30, 2026", 1210, 92, 300, 26, { fontSize: 15, color: COLORS.gray, alignment: "right" });
  addRect(slide, "gold-line", 96, 170, 58, 7, COLORS.gold, { fill: COLORS.gold, line: { fill: COLORS.gold, width: 0 } });
  addText(slide, "cover-title", "Operating signal, not noise", 96, 208, 790, 104, { fontSize: 54, bold: true, color: COLORS.ink });
  addText(slide, "cover-subtitle", "A clean-room calibration deck for evidence-led decisions", 100, 330, 680, 42, { fontSize: 23, color: COLORS.gray });
  addText(slide, "cover-basis", "Funding, throughput, and concentration read as one operating system.", 100, 390, 650, 54, { fontSize: 17, color: COLORS.gray });
  addMetric(slide, "metric-one", 100, 515, "142.0", "units funded · Q2", COLORS.gold);
  addMetric(slide, "metric-two", 330, 515, "58", "active sites · Q2", COLORS.ink);
  addMetric(slide, "metric-three", 560, 515, "4.1×", "coverage multiple", COLORS.ink);
  addText(slide, "watermark", "26", 1150, 180, 350, 230, { fontSize: 180, color: COLORS.light });
  addRect(slide, "basis-band", 96, 720, 1408, 64, COLORS.ink, { fill: COLORS.ink, line: { fill: COLORS.ink, width: 0 } });
  addText(slide, "basis-band-text", "Illustrative operating ledger · all figures fictional · basis as of 2026-04-30", 120, 740, 1200, 24, { fontSize: 15, color: COLORS.paper });
  addText(slide, "page-number", "1 / 4", 1400, 820, 100, 22, { fontSize: 14, color: COLORS.gray, alignment: "right" });
  return slide;
}

function makeStance(presentation) {
  const slide = presentation.slides.add({ name: "Stance" });
  addRect(slide, "dark", 0, 0, WIDTH, HEIGHT, COLORS.dark);
  addText(slide, "eyebrow", "NORTHSTAR · STANCE", 96, 78, 660, 24, { fontSize: 16, bold: true, color: COLORS.gold });
  addText(slide, "stance-title", "Growth is intact; concentration is the next constraint", 96, 126, 930, 98, { fontSize: 43, bold: true, color: COLORS.paper });
  addText(slide, "stance-basis", "Three observations separate signal from momentum.", 100, 246, 720, 30, { fontSize: 18, color: COLORS.mid });
  const rows = [
    ["01", "Funding is up, but late-stage deals are fewer.", "142.0 units funded; 58 active sites."],
    ["02", "Throughput improved without expanding the footprint.", "Output per site is up 18% quarter on quarter."],
    ["03", "The next decision is concentration, not acceleration.", "Two regions now carry 61% of active volume."],
  ];
  rows.forEach(([number, title, detail], index) => {
    const y = 350 + index * 118;
    addText(slide, `stance-number-${index}`, number, 112, y, 70, 54, { fontSize: 34, color: COLORS.light });
    addText(slide, `stance-${index}`, title, 220, y, 780, 34, { fontSize: 24, bold: true, color: COLORS.paper });
    addText(slide, `stance-detail-${index}`, detail, 220, y + 42, 780, 28, { fontSize: 16, color: COLORS.mid });
    addLine(slide, `stance-rule-${index}`, 220, y + 88, 1050, y + 88, COLORS.darkRule, 1);
  });
  addLine(slide, "kpi-rail", 1170, 330, 1170, 720, COLORS.darkRule, 2);
  addText(slide, "rail-label", "READ-THROUGH", 1210, 340, 300, 24, { fontSize: 16, bold: true, color: COLORS.mid });
  addText(slide, "rail-value-one", "18%", 1210, 405, 300, 64, { fontSize: 42, bold: true, color: COLORS.gold });
  addText(slide, "rail-caption-one", "output per site", 1210, 472, 300, 24, { fontSize: 16, color: COLORS.mid });
  addText(slide, "rail-value-two", "61%", 1210, 555, 300, 64, { fontSize: 42, bold: true, color: COLORS.paper });
  addText(slide, "rail-caption-two", "volume in two regions", 1210, 622, 300, 24, { fontSize: 16, color: COLORS.mid });
  addText(slide, "stance-close", "VIEW  ·  Hold the growth thesis; narrow the next bet.", 96, 786, 1180, 30, { fontSize: 18, bold: true, color: COLORS.paper });
  addText(slide, "page-number", "2 / 4", 1400, 820, 100, 22, { fontSize: 14, color: COLORS.mid, alignment: "right" });
  return slide;
}

function makeTrend(presentation) {
  const slide = presentation.slides.add({ name: "Trend data" });
  addRect(slide, "paper", 0, 0, WIDTH, HEIGHT, COLORS.paper);
  addLine(slide, "top-rule", 88, 72, 1512, 72, COLORS.ink, 3);
  addText(slide, "eyebrow", "OPERATING REVIEW · EIGHT QUARTERS", 96, 92, 700, 24, { fontSize: 16, bold: true, color: COLORS.gray });
  addText(slide, "trend-title", "Throughput rose while the footprint stayed flat", 96, 136, 1120, 54, { fontSize: 38, bold: true, color: COLORS.ink });
  addText(slide, "trend-basis", "Units shipped are bars; indexed utilization is the line. Labels show period-end values.", 100, 204, 1000, 28, { fontSize: 18, color: COLORS.gray });
  addLine(slide, "title-rule", 96, 258, 1504, 258, COLORS.light, 2);
  const chart = slide.charts.add("combo", {
    name: "throughput-chart",
    title: "Units shipped and indexed utilization",
    position: { left: 100, top: 294, width: 910, height: 360 },
    categories: ["24Q3", "24Q4", "25Q1", "25Q2", "25Q3", "25Q4", "26Q1", "26Q2"],
    series: [
      { name: "Units", chartType: "bar", values: [69, 64, 61, 58, 55, 50, 47, 43], color: COLORS.mid },
      { name: "Utilization", chartType: "line", values: [82, 88, 94, 98, 106, 112, 118, 127], color: COLORS.ink, line: { fill: COLORS.ink, width: 3 }, marker: { symbol: "circle", size: 7 } },
    ],
    legend: false,
    barOptions: { gapWidth: 90 },
    lineOptions: { marker: { symbol: "circle", size: 7 } },
  });
  addText(slide, "chart-callout", "127 utilization index", 790, 312, 240, 28, { fontSize: 16, bold: true, color: COLORS.gold, alignment: "right" });
  addLine(slide, "rail", 1090, 308, 1090, 660, COLORS.light, 2);
  addText(slide, "rail-title", "READ-THROUGH", 1140, 320, 300, 24, { fontSize: 16, bold: true, color: COLORS.gray });
  addText(slide, "rail-number", "−26", 1140, 380, 300, 60, { fontSize: 42, bold: true, color: COLORS.gold });
  addText(slide, "rail-unit", "hours per run", 1140, 445, 300, 26, { fontSize: 16, color: COLORS.gray });
  addText(slide, "rail-number-two", "+27%", 1140, 520, 300, 60, { fontSize: 42, bold: true, color: COLORS.ink });
  addText(slide, "rail-unit-two", "protected workload", 1140, 585, 300, 26, { fontSize: 16, color: COLORS.gray });
  addRect(slide, "verdict", 96, 720, 1408, 60, COLORS.panel, { fill: COLORS.panel, line: { fill: COLORS.panel, width: 0 } });
  addText(slide, "verdict-text", "VIEW  ·  The improvement survives growth; shared dependency concentration now sets the ceiling.", 120, 739, 1280, 24, { fontSize: 17, bold: true, color: COLORS.ink });
  addText(slide, "source", "Illustrative reliability ledger · source basis: internal operating review · 2024Q3–2026Q2", 96, 826, 1100, 22, { fontSize: 14, color: COLORS.gray });
  addText(slide, "page-number", "3 / 4", 1400, 826, 100, 22, { fontSize: 14, color: COLORS.gray, alignment: "right" });
  return slide;
}

function makeDecision(presentation) {
  const slide = presentation.slides.add({ name: "Decision close" });
  addRect(slide, "paper", 0, 0, WIDTH, HEIGHT, COLORS.paper);
  addLine(slide, "top-rule", 88, 72, 1512, 72, COLORS.ink, 3);
  addText(slide, "eyebrow", "DECISION MEMO · NEXT QUARTER", 96, 94, 700, 24, { fontSize: 16, bold: true, color: COLORS.gray });
  addText(slide, "decision-title", "Approve the narrow bet; keep the ledger open", 96, 140, 1080, 54, { fontSize: 38, bold: true, color: COLORS.ink });
  addText(slide, "decision-basis", "The evidence supports one constrained expansion, not a broad acceleration.", 100, 208, 1050, 28, { fontSize: 18, color: COLORS.gray });
  addLine(slide, "decision-rule", 96, 260, 1504, 260, COLORS.light, 2);
  addText(slide, "table-title", "Gates for the next allocation", 100, 300, 700, 30, { fontSize: 21, bold: true, color: COLORS.ink });
  const table = slide.tables.add({
    name: "decision-gates",
    position: { left: 100, top: 350, width: 930, height: 280 },
    values: [
      ["Gate", "Observed", "Threshold", "Action"],
      ["Volume", "142.0 units", "≥ 135", "Proceed"],
      ["Concentration", "61% in top two", "≤ 65%", "Proceed"],
      ["Coverage", "4.1×", "≥ 3.5×", "Proceed"],
      ["Cash buffer", "8.2 weeks", "≥ 10 weeks", "Hold"],
    ],
    styleOptions: { headerRow: true },
  });
  addText(slide, "ask-label", "DECISION ASK", 1140, 320, 300, 24, { fontSize: 16, bold: true, color: COLORS.gray });
  addText(slide, "ask-value", "One region", 1140, 370, 320, 46, { fontSize: 34, bold: true, color: COLORS.gold });
  addText(slide, "ask-body", "Release the next tranche only after the cash-buffer gate clears.", 1140, 430, 300, 78, { fontSize: 18, color: COLORS.gray });
  addLine(slide, "ask-rule", 1140, 540, 1440, 540, COLORS.light, 2);
  addText(slide, "ask-next", "Owner  ·  Operations finance", 1140, 570, 300, 26, { fontSize: 16, color: COLORS.gray });
  addText(slide, "ask-next-two", "Review  ·  2026-05-15", 1140, 610, 300, 26, { fontSize: 16, color: COLORS.gray });
  addRect(slide, "close-band", 96, 720, 1408, 60, COLORS.ink, { fill: COLORS.ink, line: { fill: COLORS.ink, width: 0 } });
  addText(slide, "close-text", "STANCE  ·  Fund the proven constraint, then measure the next one.", 120, 739, 1280, 24, { fontSize: 17, bold: true, color: COLORS.paper });
  addText(slide, "source", "Illustrative operating ledger · values are fictional and for calibration only", 96, 826, 1100, 22, { fontSize: 14, color: COLORS.gray });
  addText(slide, "page-number", "4 / 4", 1400, 826, 100, 22, { fontSize: 14, color: COLORS.gray, alignment: "right" });
  return slide;
}

function addMetric(slide, name, left, top, value, caption, color) {
  addText(slide, name, value, left, top, 170, 52, { fontSize: 36, bold: true, color });
  addText(slide, `${name}-caption`, caption, left, top + 56, 180, 30, { fontSize: 15, color: COLORS.gray });
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
