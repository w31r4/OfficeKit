// Run with: officekit run officekit-motion-workflow.mjs -- motion-recipes.pptx
import { writeFile } from "node:fs/promises";
import { Presentation, PresentationFile } from "office-kit";

const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });

const dataSlide = deck.slides.add({ name: "Data Rise" });
const chart = dataSlide.charts.add("bar", {
  name: "data-rise-chart",
  title: "ETF inflows",
  position: { left: 110, top: 110, width: 1060, height: 480 },
  categories: ["Mon", "Tue", "Wed", "Thu"],
  series: [{ name: "USD bn", values: [0.4, 0.8, 1.1, 1.9] }],
});
dataSlide.animations.add(chart, {
  effect: "wipe", direction: "up", chartBuild: "category-element",
  start: "onClick", durationMs: 650, staggerMs: 90,
  animateChartBackground: false,
});

const causalSlide = deck.slides.add({ name: "Causal Reveal" });
const cause = causalSlide.shapes.add({ name: "cause", geometry: "roundRect", position: { left: 100, top: 260, width: 260, height: 120 }, fill: "#DBEAFE", text: "Liquidity improves" });
const effect = causalSlide.shapes.add({ name: "effect", geometry: "roundRect", position: { left: 920, top: 260, width: 260, height: 120 }, fill: "#FEF3C7", text: "Risk assets reprice" });
const link = causalSlide.shapes.connect(cause, effect, { name: "causal-link", fromSide: "right", toSide: "left", line: { fill: "#2563EB", width: 3 }, tail: { type: "arrow" } });
causalSlide.animations.add(cause, { effect: "fade", start: "onClick" });
causalSlide.animations.add(link, { effect: "wipe", direction: "right", start: "afterPrevious" });
causalSlide.animations.add(effect, { effect: "fade", start: "afterPrevious" });

const comparisonSlide = deck.slides.add({ name: "Comparison Beat" });
const left = comparisonSlide.shapes.add({ name: "left-case", geometry: "roundRect", position: { left: 100, top: 170, width: 430, height: 280 }, fill: "#E0F2FE", text: "Scenario A" });
const right = comparisonSlide.shapes.add({ name: "right-case", geometry: "roundRect", position: { left: 750, top: 170, width: 430, height: 280 }, fill: "#FEF3C7", text: "Scenario B" });
const conclusion = comparisonSlide.shapes.add({ name: "comparison-conclusion", geometry: "textbox", position: { left: 360, top: 520, width: 560, height: 80 }, text: "Choose with evidence" });
comparisonSlide.animations.add(left, { effect: "fly", direction: "left", start: "onClick" });
comparisonSlide.animations.add(right, { effect: "fly", direction: "right", start: "withPrevious" });
comparisonSlide.animations.add(conclusion, { effect: "fade", start: "afterPrevious" });

const focusSlide = deck.slides.add({ name: "Focus Pulse" });
const riskNumber = focusSlide.shapes.add({ name: "risk-number", geometry: "textbox", position: { left: 390, top: 220, width: 500, height: 180 }, text: "80,000" });
focusSlide.animations.add(riskNumber, { phase: "emphasis", effect: "pulse", start: "afterPrevious", durationMs: 500 });

const continuitySlide = deck.slides.add({ name: "Calm Continuity" });
continuitySlide.shapes.add({ name: "continuity-title", geometry: "textbox", position: { left: 180, top: 280, width: 920, height: 100 }, text: "One calm transition" });
continuitySlide.setTransition({ effect: "fade", speed: "fast", durationMs: 450, advanceOnClick: true });

const overviewSlide = deck.slides.add({ name: "Morph Overview" });
const overviewHero = overviewSlide.shapes.add({ name: "overview-hero", geometry: "roundRect", position: { left: 120, top: 160, width: 360, height: 240 }, fill: "#FDE68A", text: "Signal" });
const detailSlide = deck.slides.add({ name: "Morph Detail" });
const detailHero = detailSlide.shapes.add({ name: "detail-hero", geometry: "roundRect", position: { left: 520, top: 100, width: 640, height: 440 }, fill: "#FDE68A", text: "Signal in detail" });
detailSlide.setMorph({ from: overviewSlide, durationMs: 800, pairs: [{ key: "hero", from: overviewHero, to: detailHero }] });

const output = process.argv[2] || "motion-recipes.pptx";
const blob = await PresentationFile.exportPptx(deck);
await writeFile(output, Buffer.from(await blob.arrayBuffer()));
console.log(deck.inspect({ kind: "animation,morph", maxChars: 20_000 }).ndjson);
console.log(`saved ${output}`);
