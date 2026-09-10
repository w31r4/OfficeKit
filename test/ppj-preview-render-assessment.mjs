import assert from "node:assert/strict";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { renderPpjToSvg } from "../src/ppj/svg-preview.mjs";
import { previewSceneFixture, nativeElement, emuFrame } from "./helpers/ppj-preview-scene-fixture.mjs";

const frame = { x: 10, y: 40, width: 200, height: 100 };
const program = {
  design: { canvas: { width: 400, height: 240 } },
  pages: [
    { id: "first", elements: [{ id: "group", type: "group", frame, elements: [
      { id: "same", type: "shape", frame, geometry: { kind: "preset", preset: "star5" }, text: "shape lost" },
      { id: "nested-source", type: "opaque", frame, nativeRef: { owner: "opaque", payload: "do-not-disclose" } },
    ] }] },
    { id: "second", elements: [{ id: "same", type: "shape", frame, geometry: { kind: "preset", preset: "star5" }, text: "also lost" }] },
  ],
};
async function render(input, pages, owners, options = {}) {
  const bytes = Buffer.from(JSON.stringify(input));
  const compiled = previewSceneFixture(input, pages, owners);
  const result = await renderPpjToSvg("in-memory.ppj", {
    load: async () => ({ path: "in-memory.ppj", program: bytes, assets: [] }), compile: async () => compiled,
    loadRaster: async () => { throw new Error("in-memory drawing must stay raster-lazy"); },
    ...options,
  });
  assert.equal(JSON.stringify(input), bytes.toString());
  return result;
}
const scenePages = preset => [
  { id: "first", elements: [nativeElement("group", "group", { ...emuFrame(),
    childWidthEmu: 2540000n, childHeightEmu: 1270000n, children: [
      nativeElement("same", "shape", { ...emuFrame(), geometry: preset, text: "shape lost" }),
      nativeElement("nested-source", "opaque", emuFrame()),
    ] })] },
  { id: "second", elements: [nativeElement("same", "shape", { ...emuFrame(), geometry: preset, text: "also lost" })] },
];
const owners = ["$.pages[0].elements[0]", "$.pages[0].elements[0].elements[0]", "$.pages[0].elements[0].elements[1]", "$.pages[1].elements[0]"];
// Keep two genuine omitted-geometry failures after rect/text painting gains
// verified profile support. Do not drop the nested identity or red-banner gate.
const result = await render(program, scenePages("star5"), owners);
assert.doesNotThrow(() => JSON.stringify(result), "the public in-memory result must not leak protobuf BigInts");
const violations = result.reliability.violations.filter((d) => d.reason === "preview.fact.shape-geometry-omitted");
assert.equal(violations.length, 2);
assert.deepEqual(new Set(violations.map((d) => d.pageId)), new Set(["first", "second"]));
assert.deepEqual(new Set(violations.map((d) => d.path)), new Set([
  "$.pages[0].elements[0].elements[0].geometry", "$.pages[1].elements[0].geometry",
]));
assert.doesNotMatch(JSON.stringify(result.assessment), /do-not-disclose/);
assert.equal(result.status, "unavailable", "unmapped preset text regions dominate opaque/partial support without clearing factual failures");
assert.equal(result.diagnostics.filter(d => d.reason === "preview.scene.paint.text-rectangle").length, 2);
assert.equal(result.pages[0].assessment.children[0].children[1].status, "opaque", "the source sibling remains independently opaque");
assert.equal(result.reliability.status, "failed");
for (const page of result.pages) {
  assert.deepEqual(page.assessment, result.assessment.children.find((child) => child.pageId === page.id));
  assert.deepEqual(page.diagnostics, page.assessment.diagnostics);
  assert.equal(page.status, page.assessment.status);
  assert.match(page.svg, /width="400" height="240" viewBox="0 0 400 240"/);
  assert.equal((page.svg.match(/data-officekit-id="same"/g) || []).length, 1);
  assert.equal((page.svg.match(/data-officekit-review="failed"/g) || []).length, 1);
}
const mappedProgram = structuredClone(program);
mappedProgram.pages[0].elements[0].elements[0].geometry.preset = "rect";
mappedProgram.pages[1].elements[0].geometry.preset = "rect";
const mapped = await render(mappedProgram, scenePages("rect"), owners);
assert.ok(!mapped.diagnostics.some(d => d.reason === "preview.fact.shape-geometry-omitted"));
assert.equal(mapped.status, "opaque", "painting the shape cannot promote the opaque sibling");
assert.ok(mapped.diagnostics.some(d => d.reason === "preview.scene.paint.text-layout"));
assert.doesNotMatch(JSON.stringify(mapped.assessment), /do-not-disclose/);
for (const page of mapped.pages) {
  assert.equal(page.reliability.status, "requires-review");
  assert.equal((page.svg.match(/data-officekit-id="same"/g) || []).length, 1);
}

// Malformed native scatter channels must not abort other pages or be
// misclassified as a missing raster dependency / missing asset. Null itself
// is now a preserved missing index, not the old canonical draw exception.
const brokenProgram = { pages: [
  { id: "broken", elements: [{ type: "chart", id: "scatter", frame, chartType: "scatter",
    data: { categories: ["A", "B"], series: [{ chartType: "scatter", values: [1, null] }] } }] },
  { id: "later", elements: [{ id: "visible", type: "text", frame, text: "later page retained" }] },
] };
const brokenPages = [
  { id: "broken", elements: [nativeElement("scatter", "chart", { ...emuFrame(), type: 6,
    series: [{ values: [1, 0], xValues: [1], missingValueIndexes: [1] }] })] },
  { id: "later", elements: [nativeElement("visible", "shape", { ...emuFrame(), geometry: "textbox", text: "later page retained" })] },
];
const brokenOwners = ["$.pages[0].elements[0]", "$.pages[1].elements[0]"];
const broken = await render(brokenProgram, brokenPages, brokenOwners);
assert.equal(broken.status, "unavailable");
assert.equal(broken.pages[0].assessment.children[0].status, "unavailable");
assert.equal(broken.pages[1].status, "partial", "a failed page must not invent failure on another page");
assert.ok(broken.diagnostics.some((d) => d.reason === "preview.scene.paint.failed" && d.path === "$.pages[0].elements[0]"));
assert.match(broken.pages[0].svg, /data-officekit-id="scatter"/);
assert.match(broken.pages[0].svg, /native drawing unavailable/);
assert.match(broken.pages[1].svg, /later page retained/);
const root = await mkdtemp(path.join(os.tmpdir(), "officekit-preview-draw-failure-"));
try {
  await assert.rejects(render(brokenProgram, brokenPages, brokenOwners, { outputDir: path.join(root, "out"),
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
  assert.ok(receipt.diagnostics.some((d) => d.reason === "preview.scene.paint.failed"));
} finally { await rm(root, { recursive: true, force: true }); }
console.log("ppj preview render assessment ok: nested identities, warning linkage and retained draw failures");
