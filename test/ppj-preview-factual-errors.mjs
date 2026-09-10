import assert from "node:assert/strict";
import registry from "../src/ppj/capability-registry.json" with { type: "json" };
import { assessPpjPreviewInput } from "../src/ppj/preview-input-assessment.mjs";
import { renderPpjToSvg } from "../src/ppj/svg-preview.mjs";
import { previewSceneFixture, nativeElement, emuFrame } from "./helpers/ppj-preview-scene-fixture.mjs";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const frame = { x: 10, y: 10, width: 300, height: 200 };
const chart = (chartType, series, extra = {}) => ({ type: "chart", chartType, frame, id: "chart", data: { categories: ["A", "B", "C"], series }, ...extra });
const samples = [
  ["shapeGeometry", { type: "shape", frame, geometry: { kind: "preset", preset: "rect" }, text: "Lost rectangle" }, ".geometry"],
  ["connector", { type: "connector", frame, from: { x: 10, y: 10 }, to: { x: 20, y: 100 } }, ".to"],
  ["transform", { type: "shape", frame: { ...frame, rotation: 45 }, geometry: { kind: "preset", preset: "rect" } }, ".frame.rotation"],
  ["visibility", { type: "text", frame, text: "must be hidden", hidden: true }, ".hidden"],
  ["groupCoordinates", { type: "group", frame, childFrame: { ...frame, x: 0 }, elements: [] }, ".childFrame"],
  ["chartChannels", chart("scatter", [{ chartType: "scatter", values: [2, 3], xValues: [10, 30] }]), ".data.series[0].xValues"],
  ["chartChannels", chart("bubble", [{ chartType: "bubble", values: [2, 3], bubbleSizes: [1, 10] }]), ".data.series[0].bubbleSizes"],
  ["chartChannels", chart("candlestick", [{ values: [2, 3], openValues: [1, 2], lowValues: [0, 1], highValues: [3, 4] }]), ".data.series[0].openValues"],
  ["chartChannels", chart("sankey", [{ values: [2, 3], sources: ["a", "b"], targets: ["c", "c"] }]), ".data.series[0].sources"],
  ["chartSeriesType", chart("column", [{ values: [1, 2] }]), ".data.series[0].chartType"],
  ["chartMissing", chart("column", [{ chartType: "column", values: [1, null, 3] }]), ".data.series[0].values[1]"],
  ["chartMissing", chart("line", [{ chartType: "line", values: [1, null, 3] }]), ".data.series[0].values[0]"],
  ["chartMissing", chart("area", [{ chartType: "area", values: [1, null, 3] }], { style: { stacking: "stream" } }), ".data.series[0].values[1]"],
  ["chartProportion", chart("pie", [{ values: [1, 9] }]), ".data"],
  ["chartHierarchy", chart("treemap", [{ values: [1, 9], parents: [null, "A"] }]), ".data.series[0].parents"],
  ["chartHierarchy", chart("sunburst", [{ values: [1, 9], levels: 2 }]), ".data.series[0].levels"],
  ["chartTotals", chart("waterfall", [{ values: [1, 9], pointRoles: ["delta", "total"] }]), ".data.series[0].pointRoles[1]"],
  ["chartScale", chart("waterfall", [{ values: [2, 3] }]), ".data.series[0].values"],
  ["chartScale", chart("combo", [{ chartType: "line", values: [1, 2], axis: "secondary" }]), ".data.series[0].axis"],
  ["chartScale", chart("line", [{ chartType: "line", values: [1, 2] }], { yAxis: { logBase: 10 } }), ".yAxis.logBase"],
  ["chartStacking", chart("area", [{ chartType: "area", values: [1, 2] }], { style: { stacking: "stream" } }), ".style.stacking"],
  ["chartSymbol", chart("column", [{ chartType: "column", values: [1, 2], symbol: { unit: 1, iconName: "person" } }]), ".data.series[0].symbol"],
];
const root = await mkdtemp(path.join(os.tmpdir(), "officekit-preview-factual-"));
let run = 0;
try { for (const [key, element, suffix] of samples) {
  const input = { pages: [{ id: "page", elements: [{ id: "target", ...element }] }] };
  const before = JSON.stringify(input), assessment = assessPpjPreviewInput(input);
  assert.equal(assessment.reliability.status, "failed", key);
  assert.ok(assessment.reliability.violations.some((d) => d.reason === registry.previewSupport.factual[key].reason && d.path === `$.pages[0].elements[0]${suffix}`), `${key} ${suffix}`);
  assert.equal(JSON.stringify(input), before);
  input.pages[0].elements[0].style = { ...input.pages[0].elements[0].style, fill: { type: "solid", color: "#FFFFFF" } };
  assert.equal(assessPpjPreviewInput(input).reliability.status, "failed", "styling cannot clear a fact failure");
  for (const program of [JSON.parse(before), input]) {
    const bytes = Buffer.from(JSON.stringify(program));
    // No verified mapping is supplied for these negative cases. A native
    // opaque node must retain every original factual gate, even with styling.
    const compiled = previewSceneFixture(program, [{ id: "page", elements: [
      nativeElement(element.id || "target", "opaque", emuFrame()),
    ] }], ["$.pages[0].elements[0]"]);
    const outputDir = path.join(root, String(run++));
    const result = await renderPpjToSvg("diagnostic.ppj", {
      outputDir, load: async () => ({ path: "diagnostic.ppj", program: bytes, assets: [] }), compile: async (_workspace, options) => {
        assert.deepEqual(options, { includeNodeMap: false, includePreviewScene: true });
        return compiled;
      },
      loadRaster: async () => ({ render: async svg => {
        assert.match(svg.toString(), /UNRELIABLE PREVIEW/);
        return Buffer.from("publication-only simulated PNG");
      } }),
    });
    assert.equal(result.reliability.status, "failed", key);
    const violation = result.reliability.violations.find((d) => d.reason === registry.previewSupport.factual[key].reason && d.path.endsWith(suffix));
    assert.ok(violation, `${key}: actual renderer must retain the original factual path`);
    const page = result.pages[0];
    assert.deepEqual(page.reliability, page.assessment.reliability);
    assert.deepEqual(page.assessment, result.assessment.children.find((child) => child.pageId === "page"));
    assert.match(page.svg, /data-officekit-review="failed"/);
    assert.match(page.svg, /UNRELIABLE PREVIEW/);
    const linked = page.svg.match(/data-officekit-diagnostic-reason="([^"]+)"/)[1];
    assert.ok(page.reliability.violations.some((d) => d.reason === linked), "visible warning links to a real violation");
    const { receipt } = result;
    assert.equal(await readFile(path.join(outputDir, receipt.pages[0].file), "utf8"), page.svg);
    assert.equal(receipt.ok, true);
    assert.equal(receipt.output.status, "complete");
    assert.equal(receipt.reliability.status, "failed", "successful publication and styling cannot clear factual failure");
    assert.deepEqual(receipt.reliability, result.reliability);
    assert.deepEqual(receipt.assessment, result.assessment);
    assert.deepEqual(receipt, JSON.parse(await readFile(path.join(outputDir, "render.json"), "utf8")));
    assert.equal(JSON.stringify(program), bytes.toString());
  }
} } finally { await rm(root, { recursive: true, force: true }); }
assert.deepEqual(new Set(samples.map(([key]) => key)), new Set(Object.keys(registry.previewSupport.factual)), "every registered factual family must be exercised");
const intactLine = assessPpjPreviewInput({ pages: [{ id: "p", elements: [chart("line", [{ chartType: "line", values: [1, 2, null, 3, 4] }])] }] });
assert.ok(!intactLine.reliability.violations.some((d) => d.reason === registry.previewSupport.factual.chartMissing.reason), "existing line gap segmentation remains a valid bounded behavior");
for (const yAxis of [{ min: 0, max: 2, reverse: false }, { logBase: { token: "scale" }, reverse: { token: "direction" } }]) {
  const limited = assessPpjPreviewInput({ pages: [{ id: "p", elements: [chart("line", [{ chartType: "line", values: [1, 2] }], { yAxis })] }] });
  assert.equal(limited.reliability.status, "requires-review", "unresolved state and matching bounds are limitations, not proven axis contradictions");
  assert.ok(limited.diagnostics.some((d) => d.path.startsWith("$.pages[0].elements[0].yAxis")));
}
console.log("ppj preview factual errors ok: each registered factual family fails closed in diagnostic evidence");
