import assert from "node:assert/strict";
import { assessPpjPreviewInput } from "../src/ppj/preview-input-assessment.mjs";
import { renderPpjToSvg } from "../src/ppj/svg-preview.mjs";

const frame = { x: 5, y: 8, width: 100, height: 60 };
const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j5aUAAAAASUVORK5CYII=", "base64");
const asset = { id: "pixel", mimeType: "image/png", data: png };
const assets = new Map([[asset.id, asset]]);
const image = { id: "picture", type: "image", frame, asset: "pixel", fit: "contain", opacity: 0.5 };
const program = (element) => ({ pages: [{ id: "page", elements: [element] }] });
const assess = (element) => assessPpjPreviewInput(program(element), { assets });
const base = assess(image);
assert.equal(base.children[0].children[0].status, "supported");
assert.equal(base.reliability.status, "passed");
const bytes = Buffer.from(JSON.stringify(program(image)));
const drawn = await renderPpjToSvg("in-memory.ppj", {
  load: async () => ({ assets: [asset] }), compile: async () => ({ programJson: bytes }),
});
assert.match(drawn.pages[0].svg, /<image data-officekit-id="picture" href="data:image\/png;base64,/u);
assert.match(drawn.pages[0].svg, /x="5" y="8" width="100" height="60" preserveAspectRatio="xMidYMid meet" opacity="0.5"/u);
assert.ok(drawn.pages[0].svg.includes(png.toString("base64")));
assert.equal(drawn.status, "supported");
assert.equal(drawn.reliability.status, "passed");
assert.deepEqual(drawn.pages[0].assessment, drawn.assessment.children[0]);
assert.doesNotMatch(drawn.pages[0].svg, /data-officekit-review=/);

const limited = assess({ ...image, crop: { left: 0.25, top: 0, right: 0, bottom: 0 } });
assert.notEqual(limited.children[0].children[0].status, "supported");
assert.ok(limited.diagnostics.some((d) => d.path.endsWith(".crop.left") && d.valueSummary === "0.25"));
const croppedInput = program({ ...image, crop: { left: 0.25, top: 0, right: 0, bottom: 0 } });
const croppedBytes = Buffer.from(JSON.stringify(croppedInput));
const cropped = await renderPpjToSvg("in-memory.ppj", {
  load: async () => ({ assets: [asset] }), compile: async () => ({ programJson: croppedBytes }),
});
assert.equal(cropped.reliability.status, "requires-review");
assert.match(cropped.pages[0].svg, /data-officekit-review="requires-review"/);
assert.match(cropped.pages[0].svg, /PREVIEW REQUIRES REVIEW/);
assert.deepEqual(cropped.canvas, drawn.canvas);
assert.equal(JSON.stringify(croppedInput), croppedBytes.toString());
const noFit = { ...image }; delete noFit.fit;
assert.ok(assess(noFit).diagnostics.some((d) => d.reason === "preview.image.default-fit.unassessed"));
assert.equal(assess({ ...image, asset: "missing" }).status, "unavailable");
assert.equal(assess({ ...image, opacity: 0 }).status, "supported");
assert.equal(assess({ ...image, frame: { ...frame, rotation: 0, flipH: false }, hidden: false }).status, "supported");
assert.notEqual(assess({ ...image, frame: { ...frame, unknownTransform: false } }).status, "supported");

const samples = [
  [{ type: "text", text: { paragraphs: [{ runs: [{ text: "A", style: { size: 24 } }] }] } }, ".text.paragraphs[0].runs[0].style.size"],
  [{ type: "shape", geometry: { kind: "custom", viewBox: frame, paths: [{ commands: [] }] } }, ".geometry.kind"],
  [{ type: "shape", geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: { token: "accent" } } } }, ".style.fill.color.token"],
  [{ ...image, mask: { kind: "preset", preset: "ellipse" } }, ".mask.preset"],
  [{ type: "table", columns: [{ width: 30 }], rows: [{ cells: [{ text: "A", rowSpan: 2 }] }] }, ".columns[0].width"],
  [{ type: "connector", from: { element: "a", anchor: "right" }, to: { element: "b", anchor: "left" } }, ".from.element"],
  [{ type: "chart", chartType: "column", data: { categories: ["A"], series: [{ id: "s", name: "S", values: [0], errorBars: { valueType: "fixed-value", value: 0 } }] } }, ".data.series[0].errorBars.value"],
  [{ type: "opaque", nativeKind: "shape", summary: "source", nativeRef: { owner: "opaque" }, previewAsset: "pixel" }, ".previewAsset"],
  [{ ...image, frame: { ...frame, rotation: 45 } }, ".frame.rotation"],
  [{ type: "group", childFrame: frame, elements: [] }, ".childFrame.width"],
];
for (const [element, suffix] of samples) {
  const input = { id: "limited", frame, ...element };
  const state = assess(input);
  assert.notEqual(state.status, "supported", suffix);
  assert.ok(state.diagnostics.some((d) => d.path === `$.pages[0].elements[0]${suffix}`), suffix);
}
const errors = assess({ id: "chart", type: "chart", frame, chartType: "column", data: { series: [{ values: [1], errorBars: { valueType: "fixed-value", value: 0 } }] } });
assert.ok(errors.diagnostics.some((d) => d.reason === "chart-error-bars-not-rendered" && d.path.endsWith(".errorBars")));
assert.ok(assess({ ...image, shadow: { blur: 10 } }).diagnostics.some((d) => d.reason === "preview.bounds.unassessed"));
console.log("ppj preview field limits ok: mapped PNG primitive and field-addressable limitations");
