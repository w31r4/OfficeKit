import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { renderPpjToSvg } from "../src/ppj/svg-preview.mjs";
import { publishPpjPreview } from "../src/ppj/preview-output.mjs";

const frame = { x: 10, y: 40, width: 200, height: 100 };
const program = {
  design: { canvas: { width: 400, height: 240 } },
  pages: [
    { id: "first", elements: [{ id: "group", type: "group", frame, elements: [
      { id: "same", type: "shape", frame, geometry: { kind: "preset", preset: "rect" }, text: "shape lost" },
      { id: "nested-source", type: "opaque", frame, nativeRef: { owner: "opaque", payload: "do-not-disclose" } },
    ] }] },
    { id: "second", elements: [{ id: "same", type: "shape", frame, geometry: { kind: "preset", preset: "rect" }, text: "also lost" }] },
  ],
};
async function render(input) {
  const bytes = Buffer.from(JSON.stringify(input));
  const result = await renderPpjToSvg("in-memory.ppj", {
    load: async () => ({ assets: [] }), compile: async () => ({ programJson: bytes }),
    loadRaster: async () => { throw new Error("in-memory drawing must stay raster-lazy"); },
  });
  assert.equal(JSON.stringify(input), bytes.toString());
  return result;
}
const result = await render(program);
const violations = result.reliability.violations.filter((d) => d.reason === "preview.fact.shape-geometry-omitted");
assert.equal(violations.length, 2);
assert.deepEqual(new Set(violations.map((d) => d.pageId)), new Set(["first", "second"]));
assert.deepEqual(new Set(violations.map((d) => d.path)), new Set([
  "$.pages[0].elements[0].elements[0].geometry", "$.pages[1].elements[0].geometry",
]));
assert.doesNotMatch(JSON.stringify(result.assessment), /do-not-disclose/);
assert.equal(result.status, "opaque", "opaque dominates partial support but never clears a factual failure");
assert.equal(result.reliability.status, "failed");
for (const page of result.pages) {
  assert.deepEqual(page.assessment, result.assessment.children.find((child) => child.pageId === page.id));
  assert.deepEqual(page.diagnostics, page.assessment.diagnostics);
  assert.equal(page.status, page.assessment.status);
  assert.match(page.svg, /width="400" height="240" viewBox="0 0 400 240"/);
  assert.equal((page.svg.match(/data-officekit-id="same"/g) || []).length, 1);
  assert.equal((page.svg.match(/data-officekit-review="failed"/g) || []).length, 1);
}

// A real known draw exception (null scatter item) must not abort other pages
// or be misclassified as a missing raster dependency / missing asset.
const broken = await render({ pages: [
  { id: "broken", elements: [{ type: "chart", id: "scatter", frame, chartType: "scatter",
    data: { categories: ["A", "B"], series: [{ chartType: "scatter", values: [1, null] }] } }] },
  { id: "later", elements: [{ id: "visible", type: "text", frame, text: "later page retained" }] },
] });
assert.equal(broken.status, "unavailable");
assert.equal(broken.pages[0].assessment.children[0].status, "unavailable");
assert.equal(broken.pages[1].status, "partial", "a failed page must not invent failure on another page");
assert.ok(broken.diagnostics.some((d) => d.reason === "preview.draw.failed" && d.path === "$.pages[0].elements[0]"));
assert.match(broken.pages[0].svg, /data-officekit-id="scatter"/);
assert.match(broken.pages[0].svg, /preview unavailable/);
assert.match(broken.pages[1].svg, /later page retained/);
const root = await mkdtemp(path.join(os.tmpdir(), "officekit-preview-draw-failure-"));
try {
  await assert.rejects(publishPpjPreview(broken, {}, { outputDir: path.join(root, "out"),
    loadRaster: async () => ({ render: async () => Buffer.from("test PNG only") }),
  }), asyncError => {
    assert.equal(asyncError.code, "preview.output.incomplete");
    assert.equal(asyncError.receipt.reliability.status, "failed");
    assert.ok(asyncError.receipt.failures.every((failure) => failure.stage === "preview"));
    assert.equal(asyncError.receipt.artifacts.length, 4);
    return true;
  });
  const receipt = JSON.parse(await readFile(path.join(root, "out", "render.json"), "utf8"));
  assert.equal(receipt.pages[0].reliability.status, "failed");
  assert.equal(receipt.pages[1].reliability.status, "requires-review");
  assert.ok(receipt.diagnostics.some((d) => d.reason === "preview.draw.failed"));
} finally { await rm(root, { recursive: true, force: true }); }
console.log("ppj preview render assessment ok: nested identities, warning linkage and retained draw failures");
