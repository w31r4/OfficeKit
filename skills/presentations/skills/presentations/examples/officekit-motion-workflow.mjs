// Run with: officekit run officekit-motion-workflow.mjs
// The task decides whether the deck is for speaking, reading, or both before
// adding one or two purposeful animations to a page.
import { writeFile } from "node:fs/promises";
import { Presentation, PresentationFile } from "office-kit";

const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const slide = deck.slides.add({ name: "motion-example" });
const chart = slide.charts.add("bar", {
  title: "ETF inflows",
  categories: ["Mon", "Tue", "Wed"],
  series: [{ name: "USD bn", values: [0.4, 0.8, 1.1] }],
});
slide.animations.add(chart, {
  effect: "wipe",
  direction: "up",
  chartBuild: "series",
  start: "onClick",
  durationMs: 650,
});

// Export the deck with the normal task-level review and output policy.
const blob = await PresentationFile.exportPptx(deck);
await writeFile("motion-example.pptx", Buffer.from(await blob.arrayBuffer()));
