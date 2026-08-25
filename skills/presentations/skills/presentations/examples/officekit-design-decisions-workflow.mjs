// Run with:
//   officekit run officekit-design-decisions-workflow.mjs -- output/design-decisions.pptx
//
// This is one deck-specific grammar expressed through five different page
// decisions. It is not a template or a universal layout helper.

import path from "node:path";
import { mkdir } from "node:fs/promises";
import { pathToFileURL } from "node:url";

import { Presentation, PresentationFile } from "office-kit";

const W = 1280;
const H = 720;
const T = {
  paper: "#F4F0E7",
  ink: "#17202A",
  quiet: "#66717D",
  accent: "#D45A32",
  secondary: "#25706A",
  pale: "#E2DED3",
  white: "#FFFDF8",
};

function frame(left, top, width, height) {
  return { left, top, width, height };
}

function svgDataUrl(svg) {
  return `data:image/svg+xml;base64,${Buffer.from(svg, "utf8").toString("base64")}`;
}

function addText(slide, name, value, position, options = {}) {
  return slide.shapes.add({
    name,
    geometry: "textbox",
    position,
    fill: "transparent",
    line: { fill: "transparent", width: 0 },
    text: value,
    textStyle: {
      fontFamily: options.fontFamily || "Aptos",
      fontSize: options.fontSize || 22,
      color: options.color || T.ink,
      bold: options.bold === true,
      italic: options.italic === true,
    },
  });
}

function addRule(slide, name, x1, y1, x2, y2, color = T.ink, width = 2) {
  return slide.connectors.add({
    name,
    start: { x: x1, y: y1 },
    end: { x: x2, y: y2 },
    line: { fill: color, width },
  });
}

function addPageLabel(slide, index, section, { folioLeft = 1160 } = {}) {
  addText(slide, `section-${index}`, section.toUpperCase(), frame(64, 36, 360, 36), {
    fontSize: 14,
    color: T.accent,
    bold: true,
  });
  addText(slide, `folio-${index}`, String(index).padStart(2, "0"), frame(folioLeft, 660, 56, 30), {
    fontSize: 12,
    color: T.quiet,
  });
}

export function addEvidenceChartPage(deck) {
  const slide = deck.slides.add({ name: "Evidence chart" });
  slide.setBackground(T.paper);
  addPageLabel(slide, 1, "evidence");
  addText(slide, "evidence-title", "Queue isolation cut P95 latency by 63%", frame(64, 88, 560, 110), {
    fontSize: 36,
    bold: true,
  });
  addText(slide, "evidence-read", "The improvement persists after traffic returns to peak load.", frame(66, 206, 520, 72), {
    fontSize: 20,
    color: T.quiet,
  });
  addRule(slide, "evidence-axis-rule", 64, 612, 1216, 612, T.ink, 1);
  slide.charts.add("bar", {
    name: "latency-evidence",
    title: "P95 latency (ms)",
    position: frame(650, 190, 530, 380),
    categories: ["Before", "Isolated", "Peak retest"],
    series: [{ name: "P95", values: [840, 310, 326], fill: T.secondary, line: { fill: T.secondary, width: 1 } }],
    yAxis: { title: "Milliseconds", min: 0, max: 900, majorUnit: 300 },
    legend: false,
  });
  addText(slide, "evidence-number", "−63%", frame(64, 340, 300, 110), {
    fontSize: 68,
    color: T.accent,
    bold: true,
  });
  addText(slide, "evidence-source", "Source: controlled load test · n=12 runs · 2026-08-24", frame(66, 622, 760, 28), {
    fontSize: 12,
    color: T.quiet,
  });
  return slide;
}

export function addImageLedPage(deck, { imageDataUrl } = {}) {
  const slide = deck.slides.add({ name: "Image-led composition" });
  slide.setBackground(T.ink);
  addPageLabel(slide, 2, "context", { folioLeft: 540 });
  const fallbackImage = svgDataUrl(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 800 720">
    <rect width="800" height="720" fill="#21303A"/>
    <path d="M0 520L220 330l150 125 170-230 260 275v220H0Z" fill="#25706A"/>
    <path d="M0 590l240-150 170 110 180-160 210 125v205H0Z" fill="#D45A32" opacity=".82"/>
    <circle cx="620" cy="145" r="72" fill="#FFFDF8" opacity=".88"/>
  </svg>`);
  slide.images.add({
    name: "context-image",
    position: frame(600, 0, 680, 720),
    dataUrl: imageDataUrl || fallbackImage,
    fit: "cover",
    alt: "Operational landscape used as the page's dominant contextual image",
  });
  slide.shapes.add({
    name: "image-caption-rail",
    geometry: "rect",
    position: frame(0, 0, 20, 720),
    fill: T.accent,
    line: { fill: T.accent, width: 0 },
  });
  addText(slide, "image-title", "Recovery starts at handoff", frame(68, 152, 470, 190), {
    fontSize: 38,
    color: T.white,
    bold: true,
  });
  addText(slide, "image-support", "A single owner and visible rollback state\nremoved the longest delay.", frame(70, 350, 430, 110), {
    fontSize: 22,
    color: T.pale,
  });
  addText(slide, "image-source", "Replace the example image with a supplied or licensed task asset.", frame(70, 628, 460, 42), {
    fontSize: 12,
    color: T.pale,
  });
  return slide;
}

export function addRelationshipDiagramPage(deck) {
  const slide = deck.slides.add({ name: "Relationship diagram" });
  slide.setBackground(T.white);
  addPageLabel(slide, 3, "system");
  addText(slide, "diagram-title", "One replay boundary prevents duplicate delivery", frame(64, 88, 880, 100), {
    fontSize: 34,
    bold: true,
  });

  const nodes = [
    ["event", "Event", 86, 282, T.paper],
    ["inbox", "Durable inbox", 382, 222, "#DCEAE7"],
    ["runner", "Agent runner", 694, 282, "#F4D8CF"],
    ["result", "Verified result", 982, 222, T.paper],
  ].map(([name, label, left, top, fill]) => slide.shapes.add({
    name: `diagram-${name}`,
    geometry: "rect",
    position: frame(left, top, 210, 104),
    fill,
    line: { fill: T.ink, width: 1.5 },
    text: label,
    textStyle: { fontFamily: "Aptos", fontSize: 22, color: T.ink, bold: true },
  }));

  for (let index = 0; index < nodes.length - 1; index += 1) {
    slide.shapes.connect(nodes[index], nodes[index + 1], {
      name: `diagram-link-${index + 1}`,
      fromSide: "right",
      toSide: "left",
      line: { fill: index === 1 ? T.accent : T.secondary, width: 3 },
      tail: { type: "arrow" },
    });
  }
  addText(slide, "diagram-boundary", "Idempotency is checked here", frame(385, 410, 300, 42), {
    fontSize: 16,
    color: T.accent,
    bold: true,
  });
  addRule(slide, "diagram-boundary-rule", 486, 394, 486, 338, T.accent, 2);
  addText(slide, "diagram-consequence", "Retries wake the runner; they do not create a second business event.", frame(382, 510, 610, 90), {
    fontSize: 24,
    bold: true,
  });
  return slide;
}

export function addAsymmetricEditorialPage(deck) {
  const slide = deck.slides.add({ name: "Asymmetric editorial" });
  slide.setBackground(T.paper);
  addPageLabel(slide, 4, "decision");
  addText(slide, "editorial-kicker", "THE DECISION", frame(64, 122, 260, 30), {
    fontSize: 14,
    color: T.accent,
    bold: true,
  });
  addText(slide, "editorial-title", "Ship the bounded path now.", frame(64, 170, 610, 130), {
    fontSize: 48,
    bold: true,
  });
  addText(slide, "editorial-title-2", "Move residual risk into one explicit gate.", frame(64, 320, 700, 108), {
    fontSize: 40,
    color: T.secondary,
    bold: true,
  });
  addRule(slide, "editorial-rule", 798, 106, 798, 602, T.accent, 5);
  addText(slide, "editorial-proof", "12/12 recovery runs passed", frame(850, 184, 330, 86), {
    fontSize: 30,
    bold: true,
  });
  addText(slide, "editorial-condition", "Condition\nExpand only after the rollback\nsignal remains visible for 24 hours.", frame(850, 322, 330, 156), {
    fontSize: 18,
    color: T.quiet,
  });
  addText(slide, "editorial-action", "Owner: Runtime team · Gate: 10% traffic", frame(64, 600, 700, 36), {
    fontSize: 16,
    color: T.quiet,
  });
  return slide;
}

export function addRestrainedMotifPage(deck) {
  const slide = deck.slides.add({ name: "Restrained motif" });
  slide.setBackground(T.white);
  addPageLabel(slide, 5, "close");
  addText(slide, "motif-title", "A motif should orient the reader, then get out of the way", frame(64, 92, 980, 110), {
    fontSize: 34,
    bold: true,
  });
  addRule(slide, "motif-baseline", 64, 582, 1216, 582, T.pale, 2);
  const xPositions = [108, 358, 608];
  const labels = ["Claim", "Evidence", "Action"];
  for (let index = 0; index < xPositions.length; index += 1) {
    const left = xPositions[index];
    slide.shapes.add({
      name: `motif-marker-${index + 1}`,
      geometry: "rect",
      position: frame(left, 248, 14, 182),
      fill: index === 1 ? T.accent : T.secondary,
      line: { fill: "transparent", width: 0 },
    });
    addText(slide, `motif-label-${index + 1}`, labels[index], frame(left + 32, 248, 180, 44), {
      fontSize: 24,
      bold: true,
    });
  }
  addText(slide, "motif-explanation", "The same narrow rule marks reading stages. It never becomes a container, button, or decoration cloud.", frame(850, 248, 300, 180), {
    fontSize: 22,
    color: T.quiet,
  });
  addText(slide, "motif-close", "Repeat the rule only where it carries this meaning.", frame(64, 612, 720, 42), {
    fontSize: 18,
    color: T.secondary,
    bold: true,
  });
  return slide;
}

export function createDesignDecisionDeck({ imageDataUrl } = {}) {
  const deck = Presentation.create({ slideSize: { width: W, height: H } });
  addEvidenceChartPage(deck);
  addImageLedPage(deck, { imageDataUrl });
  addRelationshipDiagramPage(deck);
  addAsymmetricEditorialPage(deck);
  addRestrainedMotifPage(deck);
  return deck;
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const output = path.resolve(process.argv[2] || "output/design-decisions.pptx");
  await mkdir(path.dirname(output), { recursive: true });
  const deck = createDesignDecisionDeck();
  const verification = deck.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Design-decision example failed verification: ${verification.ndjson}`);
  await (await PresentationFile.exportPptx(deck)).save(output);
  console.log(JSON.stringify({ output, slides: deck.slides.items.length, verification: "passed" }));
}
