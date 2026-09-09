// Explicit integration against a checked-in-build-command output. This does not
// replace the installed package. Internal SVG foundations are exercised below;
// production scene routing/publication (tasks 3.3/4.x/5.2) remain separate.
import assert from "node:assert/strict";
import path from "node:path";
import { registerHooks } from "node:module";
import JSZip from "jszip";
import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import { CodecRequestSchema, CodecResponseSchema } from "../src/generated/office_kit/artifact/v1/office_artifact_pb.js";
import { loadOfficeKitNativeDescriptor, startOfficeKitNativeClient } from "../src/codecs/office-kit-native-client.mjs";
import { readPpjPreviewScene } from "../src/ppj/preview-scene.mjs";
import { createPpjSceneView } from "../src/ppj/preview-scene-view.mjs";
import { paintPpjSceneSvg, nativePathData } from "../src/ppj/preview-scene-svg.mjs";
import { mkdtemp, writeFile } from "node:fs/promises";
import os from "node:os";

assert.ok(process.argv[2], "Pass the directory produced by npm run build:office-kit -- --output <new-directory>.");
const packageJsonPath = path.resolve(process.argv[2], "package.json");
const descriptors = await Promise.all(["office", "ppj"].map(profile => loadOfficeKitNativeDescriptor({ packageJsonPath, profile })));
const started = [];
globalThis[Symbol.for("officekit.preview.native.test")] = async (options) => {
  const client = await startOfficeKitNativeClient({ ...options, packageJsonPath });
  started.push(options.profile);
  return client;
};
const stub = `export const OFFICE_KIT_NATIVE_TRANSPORT_VERSION=2;
  export const startOfficeKitNativeClient=(options)=>globalThis[Symbol.for("officekit.preview.native.test")](options);`;
const hook = registerHooks({ resolve(specifier, context, resolve) {
  if (specifier === "./office-kit-native-client.mjs" && context.parentURL.endsWith("/src/codecs/office-kit-runtime.mjs"))
    return { url: `data:text/javascript,${encodeURIComponent(stub)}`, shortCircuit: true };
  return resolve(specifier, context);
} });
// The packaged Office profile is intentionally Word/Excel-only. The general
// CodecProtocol library entrypoint is covered by the native test suite, not by
// routing presentation operations into officekit-codec. Exercise the generated
// full wire reader independently against the real PPJ executable here.
try {
  const artifacts = await mkdtemp(path.join(os.tmpdir(), "officekit-native-scene-paint-"));
  console.log(`Native scene integration artifacts: ${artifacts}`);
  const relationFailures = [];
  let directedAnchorCases = 0;
  let customArcPath = false;
  const { default: sharp } = await import("sharp");
  async function savePaint(name, receipt) {
    const painted = paintPpjSceneSvg(receipt);
    await writeFile(path.join(artifacts, `${name}.diagnostics.json`), JSON.stringify(painted.diagnostics, null, 2), { flag: "wx" });
    for (const [i, page] of painted.pages.entries()) {
      await writeFile(path.join(artifacts, `${name}-${i}.svg`), page.svg, { flag: "wx" });
      const png = await sharp(Buffer.from(page.svg)).png().toBuffer();
      await writeFile(path.join(artifacts, `${name}-${i}.png`), png, { flag: "wx" });
      assert.equal((await sharp(png).metadata()).width, painted.canvas.width);
    }
    return painted;
  }
  const { loadPpjWorkspace, compilePpjWorkspace, validatePpjWorkspace, sha256 } = await import("../src/ppj/workspace.mjs");
  const { projectPptxToPpj } = await import("../src/ppj/native.mjs");
  const { invokeOfficeKitLazy } = await import("../src/codecs/office-kit-runtime.mjs");
  async function withoutAuthoredSnapshot(file) {
    const archive = await JSZip.loadAsync(file);
    for (const name of Object.keys(archive.files)) if (name.startsWith("officeKit/")) archive.remove(name);
    archive.file("_rels/.rels", (await archive.file("_rels/.rels").async("string")).replace(
      /<Relationship\b(?=[^>]*\bType="https:\/\/schemas\.officekit\.dev\/relationships\/presentation-program")[^>]*(?:\/>|>[\s\S]*?<\/Relationship>)/g, ""));
    archive.file("[Content_Types].xml", (await archive.file("[Content_Types].xml").async("string")).replace(
      /<Override\b(?=[^>]*\bPartName="\/officeKit\/)[^>]*(?:\/>|>[\s\S]*?<\/Override>)/g, ""));
    return archive.generateAsync({ type: "uint8array" });
  }
  const fullWireCompile = async (workspace, { includePreviewScene = true, validationOnly = false } = {}) => {
    const request = create(CodecRequestSchema, {
    protocolVersion: 2, operation: 11, family: 2,
    presentationProgram: { programJson: workspace.program, includePreviewScene, validationOnly, includeNodeMap: true,
      assets: workspace.assets.map(asset => ({ id: asset.id, contentType: asset.mimeType, sha256: asset.sha256, data: asset.data })) },
    });
    // Independent comparisons can be separated by a long raster/edit suite.
    // Own the client per comparison; never reuse a handle after idle retirement
    // or change production idle policy to keep this test's handle alive.
    const client = await startOfficeKitNativeClient({ packageJsonPath, profile: "ppj" });
    try {
      const response = fromBinary(CodecResponseSchema,
        await client.invoke(toBinary(CodecRequestSchema, request), workspace.source), { recursionLimit: 136 });
      assert.equal(response.ok, true, JSON.stringify(response.diagnostics));
      return { file: response.file, program: response.presentationProgram,
        ...(includePreviewScene ? { previewScene: readPpjPreviewScene(response.presentationProgram, response.file) } : {}) };
    } finally { await client.retire(); }
  };
  await assert.rejects(invokeOfficeKitLazy(() => ({ protocolVersion: 2, operation: 11, family: 2,
    presentationProgram: { programJson: new TextEncoder().encode("{}"), includePreviewScene: true },
  })), error => error.code === "unsupported_operation");
  let authoredAssets;
  for (const fixture of ["examples/ppj/minimum.ppj", "test/fixtures/presentation/evidence-ledger-canonical.ppj"]) {
    const workspace = await loadPpjWorkspace(fixture);
    const before = workspace.program.slice();
    const ordinary = await compilePpjWorkspace(workspace);
    const checked = await validatePpjWorkspace(workspace);
    assert.equal(Object.hasOwn(ordinary, "previewScene"), false);
    assert.equal(Object.hasOwn(checked, "previewScene"), false);
    const ppj = await compilePpjWorkspace(workspace, { includePreviewScene: true });
    const view = createPpjSceneView(ppj);
    assert.equal(view.scene, ppj.previewScene);
    assert.equal(view.pages.length, ppj.previewScene.presentation.slides.length);
    assert.equal(view.pages[0].nodes[0].native, ppj.previewScene.presentation.slides[0].elements[0].content.value);
    assert.equal(view.paintAssessment, "unassessed");
    assert.deepEqual(view.diagnostics, []);
    if (fixture.includes("minimum")) assert.equal(view.pages[0].nodes[0].frame.x, 48);
    const painted = await savePaint(fixture.includes("minimum") ? "minimum" : "canonical", ppj);
    assert.match(painted.pages[0].svg, /data-officekit-native-id=/);
    const full = await fullWireCompile(workspace);
    const fullOrdinary = await fullWireCompile(workspace, { includePreviewScene: false });
    const fullChecked = await fullWireCompile(workspace, { includePreviewScene: false, validationOnly: true });
    assert.deepEqual(ppj.file, ordinary.file);
    assert.deepEqual(Uint8Array.from(full.file), ordinary.file);
    assert.deepEqual(Uint8Array.from(fullOrdinary.file), ordinary.file);
    assert.equal(fullOrdinary.program.previewScene, undefined);
    assert.equal(fullChecked.program.previewScene, undefined);
    assert.deepEqual(full.previewScene, ppj.previewScene);
    assert.deepEqual(workspace.program, before);
    assert.equal(ppj.previewScene.origin, 1);
    assert.equal(ppj.previewScene.candidateSha256, sha256(ppj.file));
    if (fixture.includes("canonical")) {
      assert.ok(ppj.previewScene.bindings.some(binding => binding.componentId && binding.attribution === 2));
      assert.ok(ppj.previewScene.assets.length);
      authoredAssets = workspace;
    }
  }
  const sourceWorkspace = await loadPpjWorkspace("examples/ppj/minimum.ppj");
  // One bounded component/explicit pair, with the same genuine zero/missing
  // chart edge in both. This is not all of task 5.1's required equivalences.
  const pairBase = JSON.parse(new TextDecoder().decode(sourceWorkspace.program));
  const primitive = { id: "tile", type: "shape", frame: { x: 10, y: 10, width: 60, height: 40 },
    geometry: { kind: "custom", viewBox: { x: 0, y: 0, width: 100, height: 100 }, paths: [{ fill: true, stroke: false,
      commands: [{ op: "moveTo", x: 50, y: 0 }, { op: "arcTo", radiusX: 50, radiusY: 25, startAngle: 45, sweepAngle: 90 },
        { op: "lineTo", x: 0, y: 100 }, { op: "close" }] }] },
    style: { fill: { type: "solid", color: "#CC5500" } } };
  const zeroLabel = { id: "zero", type: "text", frame: { x: 10, y: 60, width: 80, height: 30 }, text: "0" };
  const missingLine = { id: "missing-line", type: "chart", chartType: "line", frame: { x: 500, y: 100, width: 300, height: 200 },
    data: { categories: ["A", "B", "C"], series: [{ id: "values", name: "Values", values: [1, null, 0] }] } };
  const componentProgram = structuredClone(pairBase);
  componentProgram.components = [{ id: "tile-component", frame: { x: 0, y: 0, width: 200, height: 100 }, elements: [primitive, zeroLabel] }];
  componentProgram.pages[0].elements = [{ id: "instance", type: "component", component: "tile-component", frame: { x: 100, y: 100, width: 200, height: 100 } }, missingLine];
  const explicitProgram = structuredClone(pairBase);
  explicitProgram.pages[0].elements = [primitive, zeroLabel].map(e => ({ ...structuredClone(e), frame: { ...e.frame, x: e.frame.x + 100, y: e.frame.y + 100 } }));
  explicitProgram.pages[0].elements.push(missingLine);
  const pair = [];
  for (const [i, program] of [componentProgram, explicitProgram].entries()) {
    const compiled = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) }, { includePreviewScene: true });
    const painted = await savePaint(`pair-${i}`, compiled);
    const scene = createPpjSceneView(compiled);
    const leaves = scene.pages[0].nodes.filter(n => n.kind === "shape");
    assert.equal(leaves.length, 2);
    assert.deepEqual(leaves[0].frame, { x: 110, y: 110, width: 60, height: 40 });
    assert.equal(leaves[0].native.fillRgb, "CC5500");
    assert.equal(leaves[1].native.text, "0");
    const chart = scene.pages[0].nodes.find(n => n.kind === "chart").native;
    assert.deepEqual(chart.series[0].values, [1, 0, 0]);
    assert.deepEqual(chart.series[0].missingValueIndexes, [1]);
    assert.match(painted.pages[0].svg, /data-officekit-chart="line"/);
    assert.match(painted.pages[0].svg, /data-officekit-point="0" data-officekit-value="1"/);
    assert.match(painted.pages[0].svg, /data-officekit-point="2" data-officekit-value="0"/);
    assert.doesNotMatch(painted.pages[0].svg, /data-officekit-line-segment=/);
    assert.equal((painted.pages[0].svg.match(/data-officekit-review-point="isolated"/g) || []).length, 2);
    assert.match(painted.pages[0].svg, /data-officekit-path="0" d="M 140 110 A 30 10 0 0 1 113\.16718427[0-9]+ 110 L 110 150 Z"/);
    customArcPath = true;
    assert.match(painted.pages[0].svg, />0<\/tspan>/);
    const pixel = await sharp(Buffer.from(painted.pages[0].svg)).extract({ left: 118, top: 130, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
    assert.deepEqual([...pixel], [204, 85, 0]);
    pair.push(leaves.map(n => ({ frame: n.frame, fill: n.native.fillRgb, text: n.native.text })));
  }
  assert.deepEqual(pair[0], pair[1]);
  const nativeLineInput = { id: "native-line", type: "chart", chartType: "line", title: "Missing is not zero",
    frame: { x: 500, y: 330, width: 400, height: 160 }, yAxis: { min: 0, max: 5 },
    data: { categories: ["A", "Missing", "Zero", "D", "E"], series: [{ id: "observed", name: "Observed",
      values: [2, null, 0, 4, 5], stroke: { color: "#047857", width: 2 },
      marker: { symbol: "circle", size: 8, fill: "#047857", stroke: { color: "#114477", width: 2 } } },
      { id: "other", name: "Other series", values: [5, 4, null, 0, 1], stroke: { color: "#B45309", width: 2 }, marker: "none" }] } };
  async function assertNativeLine(receipt, painted, middle = 0) {
    const native = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "chart").native;
    assert.deepEqual(native.series[0].values, [2, 0, middle, 4, 5]);
    assert.deepEqual(native.series[0].missingValueIndexes, [1]);
    assert.equal(native.series[0].marker.size, 8);
    assert.equal(native.series[0].marker.line.color.source.value, "114477");
    assert.equal(native.series[0].marker.line.widthPoints, 2);
    assert.equal(native.series.length, 2);
    assert.deepEqual(native.series[1].values, [5, 4, 0, 0, 1]);
    assert.deepEqual(native.series[1].missingValueIndexes, [2]);
    const svg = painted.pages[0].svg;
    assert.match(svg, /data-officekit-chart="line"/);
    assert.match(svg, /data-officekit-missing-point="1"/);
    const firstSeriesSvg = svg.slice(svg.indexOf('data-officekit-series="0"'), svg.indexOf('data-officekit-series="1"'));
    const secondSeriesSvg = svg.slice(svg.indexOf('data-officekit-series="1"'));
    assert.doesNotMatch(firstSeriesSvg, /data-officekit-point="1"|data-officekit-line-segment="0:/);
    assert.match(secondSeriesSvg, /data-officekit-line-segment="0:1" d="M 540 354 L 620 376.4"[^>]*stroke="#B45309"/);
    assert.match(secondSeriesSvg, /data-officekit-line-segment="3:4" d="M 780 466 L 860 443.6"/);
    assert.match(secondSeriesSvg, /data-officekit-missing-point="2"/);
    assert.doesNotMatch(secondSeriesSvg, /data-officekit-point="2"/);
    assert.ok(svg.includes(`data-officekit-point="2" data-officekit-value="${middle}"`));
    assert.ok(svg.includes(`data-officekit-line-segment="2:4" d="M 700 ${354 + (1 - middle / 5) * 112} L 780 376.4 L 860 354"`));
    assert.ok(!painted.diagnostics.some(d => ["preview.scene.paint.chart-semantics", "preview.scene.paint.failed"].includes(d.reason)));
    // Actual marker interior, and empty space where joining across the gap
    // would draw a fictitious segment. No Office or mocked raster backend.
    // Left/top/right plot-edge markers remain whole, while the connecting
    // paths are clipped. These pixels were outside the old plot viewport.
    for (const [left, top, rgb] of [[780, 376, [4, 120, 87]], [620, 444, [255, 255, 255]],
      [538, 421, [4, 120, 87]], [536, 421, [17, 68, 119]], [861, 354, [4, 120, 87]], [860, 350, [17, 68, 119]],
      ...(middle === 0 ? [[700, 467, [4, 120, 87]]] : [])]) {
      const pixel = await sharp(Buffer.from(svg)).extract({ left, top, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
      assert.deepEqual([...pixel], rgb);
    }
  }
  const nativeLineProgram = structuredClone(pairBase);
  nativeLineProgram.pages[0].elements = [structuredClone(nativeLineInput)];
  const nativeLineResult = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(nativeLineProgram)) }, { includePreviewScene: true });
  await assertNativeLine(nativeLineResult, await savePaint("native-line", nativeLineResult));
  const nativeBars = [];
  for (const type of ["column", "bar"]) {
    const program = structuredClone(pairBase);
    program.pages[0].elements = [{ id: "bars", type: "chart", chartType: type, title: "Missing, zero and negative",
      frame: { x: 60, y: 80, width: 450, height: 300 }, yAxis: { min: -4, max: 8 },
      style: { gapWidth: 100, overlap: 0, stacking: "none" },
      data: { categories: ["Positive", "Unknown", "Zero", "Negative"], series: [
        { id: "first", name: "First", values: [4, null, 0, -4], fill: { type: "solid", color: "#CC2200" },
          pointStyles: [{ index: 3, fill: { type: "solid", color: "#00AA44" } }] },
        { id: "second", name: "Second", values: [8, 2, null, -2], fill: { type: "solid", color: "#0044CC" } },
      ] } }];
    const compileBar = value => compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(value)) }, { includePreviewScene: true });
    async function checkBar(receipt, name, { first = 4, gap = 100, overlap = 0, reverse = false } = {}) {
      const native = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "chart").native;
      assert.equal(native.type, 1); assert.equal(native.barDirection, type); assert.equal(native.grouping, "none");
      assert.deepEqual(native.series[0].values, [first, 0, 0, -4]); assert.deepEqual(native.series[0].missingValueIndexes, [1]);
      assert.deepEqual(native.series[1].values, [8, 2, 0, -2]); assert.deepEqual(native.series[1].missingValueIndexes, [2]);
      assert.equal(native.gapWidth, gap === null ? undefined : gap); assert.equal(native.overlap, overlap === null ? undefined : overlap);
      assert.equal(native.xAxis?.reverse ?? false, reverse); assert.equal(native.yAxis?.reverse ?? false, reverse);
      const painted = await savePaint(name, receipt), svg = painted.pages[0].svg;
      assert.ok(!painted.diagnostics.some(d => ["preview.scene.paint.chart-semantics", "preview.scene.paint.failed"].includes(d.reason)));
      assert.match(svg, new RegExp(`data-officekit-chart="${type}"`));
      assert.equal((svg.match(/data-officekit-missing-point=/g) || []).length, 2);
      assert.equal((svg.match(/data-officekit-review-point="zero"/g) || []).length, 1);
      const actual = [...svg.matchAll(/<rect data-officekit-bar="(\d+)" x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"[^>]*fill="([^"]+)"/g)]
        .map(m => ({ index: +m[1], rect: m.slice(2, 6).map(Number), color: m[6] }));
      const band = type === "column" ? 90 : 52.5, width = band / (2 - (overlap ?? 0) / 100 + (gap ?? 150) / 100);
      const stride = width * (1 - (overlap ?? 0) / 100), inset = (band - width - stride) / 2;
      const rectangle = (si, index, value) => {
        const offset = index * band + inset + si * stride;
        const rect = type === "column" ? [105 + offset, value > 0 ? 265 - value * 17.5 : 265, width, Math.abs(value) * 17.5]
          : [value > 0 ? 225 : 225 + value * 30, 335 - offset - width, Math.abs(value) * 30, width];
        return reverse ? [570 - rect[0] - rect[2], 460 - rect[1] - rect[3], rect[2], rect[3]] : rect;
      };
      const expected = [[0, 0, first, "#CC2200"], [0, 3, -4, "#00AA44"], [1, 0, 8, "#0044CC"], [1, 1, 2, "#0044CC"], [1, 3, -2, "#0044CC"]];
      assert.equal(actual.length, expected.length);
      for (const [j, [si, index, value, color]] of expected.entries()) {
        const rect = rectangle(si, index, value);
        assert.equal(actual[j].index, index); assert.equal(actual[j].color.toUpperCase(), color);
        actual[j].rect.forEach((v, k) => assert.ok(Math.abs(v - rect[k]) < 1e-9, `${name} bar geometry ${j}:${k}`));
        const pixel = await sharp(Buffer.from(svg)).extract({ left: Math.floor(rect[0] + rect[2] / 2), top: Math.floor(rect[1] + rect[3] / 2), width: 1, height: 1 }).removeAlpha().raw().toBuffer();
        assert.deepEqual([...pixel], color.slice(1).match(/../g).map(v => parseInt(v, 16)), `${name} bar ${si}:${index} pixel`);
      }
      // Sample the absent first-series bar's reserved slot, independently of
      // the SVG output; the other series must not shift into the missing slot.
      const absent = rectangle(0, 1, 4);
      const empty = await sharp(Buffer.from(svg)).extract({ left: Math.floor(absent[0] + absent[2] / 2), top: Math.floor(absent[1] + absent[3] / 2), width: 1, height: 1 }).removeAlpha().raw().toBuffer();
      assert.deepEqual([...empty], [255, 255, 255]);
    }
    const authored = await compileBar(program);
    await checkBar(authored, `${type}-authored`);
    const reversed = structuredClone(program);
    reversed.pages[0].elements[0].xAxis = { reverse: true };
    reversed.pages[0].elements[0].yAxis.reverse = true;
    await checkBar(await compileBar(reversed), `${type}-reversed`, { reverse: true });
    const source = await withoutAuthoredSnapshot(authored.file), before = source.slice();
    const projected = await projectPptxToPpj(source, { sourceUri: `${type}-source.pptx`, assetRootUri: "assets" });
    assert.equal(projected.sourceBound, true);
    const sourceInput = { program: projected.programJson, source, assets: projected.assets };
    const noop = await compilePpjWorkspace(sourceInput, { includePreviewScene: true });
    assert.deepEqual(noop.file, source); await checkBar(noop, `${type}-source-noop`);
    const edits = [];
    for (const [name, expected, edit] of [
      ["value", { first: 6 }, c => { c.data.series[0].values[0] = 6; }],
      ["gap-zero", { gap: 0 }, c => { c.style.gapWidth = 0; }],
      ["overlap-negative", { overlap: -100 }, c => { c.style.overlap = -100; }],
      ["gap-delete", { gap: null }, c => { delete c.style.gapWidth; }],
      ["overlap-delete", { overlap: null }, c => { delete c.style.overlap; }],
    ]) {
      const changed = JSON.parse(new TextDecoder().decode(projected.programJson));
      edit(changed.pages[0].elements.find(e => e.type === "chart"));
      const candidate = await compilePpjWorkspace({ ...sourceInput, program: Buffer.from(JSON.stringify(changed)) }, { includePreviewScene: true });
      await checkBar(candidate, `${type}-source-${name}`, expected);
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: `${type}-${name}.pptx`, assetRootUri: "assets" });
      const result = JSON.parse(new TextDecoder().decode(fresh.programJson)).pages[0].elements.find(e => e.type === "chart");
      assert.equal(result.chartType, type);
      assert.deepEqual(result.data.series[0].values, [expected.first ?? 4, null, 0, -4]);
      assert.deepEqual(result.data.series[1].values, [8, 2, null, -2]);
      assert.equal(result.style.gapWidth, expected.gap === null ? undefined : expected.gap ?? 100);
      assert.equal(result.style.overlap, expected.overlap === null ? undefined : expected.overlap ?? 0);
      const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file);
      assert.deepEqual(Object.keys(newZip.files).sort(), Object.keys(oldZip.files).sort());
      const parts = [];
      for (const part of Object.keys(oldZip.files)) if (!oldZip.files[part].dir &&
        !Buffer.from(await oldZip.file(part).async("uint8array")).equals(Buffer.from(await newZip.file(part).async("uint8array")))) parts.push(part);
      assert.deepEqual(parts.sort(), ["ppt/slides/charts/chart1.xml"]);
      assert.ok(!Object.keys(oldZip.files).some(p => p.endsWith(".xlsx")));
      assert.deepEqual(source, before);
      edits.push({ name, changedParts: parts, sourceSha256: sha256(source), candidateSha256: sha256(candidate.file), pixels: true, reprojection: true });
    }
    nativeBars.push({ type, authored: true, reversedAxes: true, sourceNoop: true, negativeZeroMissing: true, seriesSlotsAndPointPaint: true, edits, workbook: "not present in literal-data fixture" });
  }
  const nativeStacks = [];
  for (const type of ["column", "bar"]) for (const grouping of ["stacked", "percent-stacked"]) {
    const percent = grouping === "percent-stacked", suffix = percent ? "percent" : "stack";
    const lastFirst = percent ? 4 : -4, lastSecond = percent ? 2 : -2;
    const program = structuredClone(pairBase);
    program.pages[0].elements = [{ id: "stack", type: "chart", chartType: type, title: percent ? "Complete-category shares" : "Signed stacks with unknown data",
      frame: { x: 60, y: 80, width: 450, height: 300 }, style: { stacking: grouping, gapWidth: 100, overlap: 100 },
      data: { categories: ["Positive", "Unknown", "Zero", percent ? "Other" : "Negative"], series: [
        { id: "first", name: "First", values: [4, null, 0, lastFirst], fill: { type: "solid", color: "#CC2200" } },
        { id: "second", name: "Second", values: [8, 2, 0, lastSecond], fill: { type: "solid", color: "#0044CC" } },
      ] } }];
    async function checkStack(receipt, name, first = 4) {
      const native = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "chart").native;
      assert.equal(native.grouping, grouping); assert.equal(native.barDirection, type); assert.equal(native.overlap, 100);
      assert.deepEqual(native.series[0].values, [first, 0, 0, lastFirst]); assert.deepEqual(native.series[0].missingValueIndexes, [1]);
      assert.deepEqual(native.series[1].values, [8, 2, 0, lastSecond]); assert.deepEqual(native.series[1].missingValueIndexes, []);
      const painted = await savePaint(name, receipt), svg = painted.pages[0].svg;
      assert.equal(painted.reliability.status, "requires-review");
      assert.match(svg, /data-officekit-incomplete-stack="1"/);
      assert.match(svg, /data-officekit-point="1" data-officekit-value="2" data-officekit-stack-position="unknown"/);
      assert.equal((svg.match(/data-officekit-review-point="zero"/g) || []).length, percent ? 0 : 2);
      if (percent) {
        assert.match(svg, /data-officekit-zero-total-stack="2"/);
        assert.match(svg, /data-officekit-value-tick="1"[^>]*>100%<\/text>/);
        const fractions = [...svg.matchAll(/data-officekit-fraction="([^"]+)"/g)].map(m => Number(m[1]));
        [first / (first + 8), 4 / 6, 8 / (first + 8), 2 / 6].forEach((v, i) => assert.ok(Math.abs(fractions[i] - v) < 1e-12));
        assert.equal(fractions.length, 4);
      } else {
        assert.match(svg, new RegExp(`data-officekit-value="8" data-officekit-baseline="${first}" data-officekit-stack-end="${first + 8}"`));
        assert.match(svg, /data-officekit-value="-2" data-officekit-baseline="-4" data-officekit-stack-end="-6"/);
      }
      const actual = [...svg.matchAll(/<rect data-officekit-bar="(\d+)" x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"/g)].map(m => m.slice(1).map(Number));
      const expected = percent
        ? [[0, 0, first / (first + 8), "CC2200"], [3, 0, 4 / 6, "CC2200"], [0, first / (first + 8), 1, "0044CC"], [3, 4 / 6, 1, "0044CC"]]
        : [[0, 0, first, "CC2200"], [3, 0, -4, "CC2200"], [0, first, first + 8, "0044CC"], [3, -4, -6, "0044CC"]];
      assert.equal(actual.length, expected.length);
      for (const [j, [i, start, end, color]] of expected.entries()) {
        const scale = (type === "column" ? 210 : 360) / (percent ? 1 : first + 14), width = type === "column" ? 45 : 26.25, offset = percent ? 0 : 6;
        const rect = type === "column"
          ? [127.5 + i * 90, 335 - (Math.max(start, end) + offset) * scale, width, Math.abs(end - start) * scale]
          : [105 + (Math.min(start, end) + offset) * scale, 295.625 - i * 52.5, Math.abs(end - start) * scale, width];
        actual[j].forEach((v, k) => assert.ok(Math.abs(v - [i, ...rect][k]) < 1e-9, `${name} cumulative geometry ${j}:${k}`));
        const pixel = await sharp(Buffer.from(svg)).extract({ left: Math.floor(rect[0] + rect[2] / 2), top: Math.floor(rect[1] + rect[3] / 2), width: 1, height: 1 }).removeAlpha().raw().toBuffer();
        assert.deepEqual([...pixel], color.match(/../g).map(v => parseInt(v, 16)));
      }
    }
    const authored = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) }, { includePreviewScene: true });
    await checkStack(authored, `${type}-${suffix}-authored`);
    const source = await withoutAuthoredSnapshot(authored.file), original = source.slice();
    const projected = await projectPptxToPpj(source, { sourceUri: `${type}-${suffix}.pptx`, assetRootUri: "assets" });
    const input = { program: projected.programJson, source, assets: projected.assets };
    const noop = await compilePpjWorkspace(input, { includePreviewScene: true });
    assert.deepEqual(noop.file, source); await checkStack(noop, `${type}-${suffix}-noop`);
    const changed = JSON.parse(new TextDecoder().decode(projected.programJson));
    changed.pages[0].elements.find(e => e.type === "chart").data.series[0].values[0] = 6;
    const candidate = await compilePpjWorkspace({ ...input, program: Buffer.from(JSON.stringify(changed)) }, { includePreviewScene: true });
    await checkStack(candidate, `${type}-${suffix}-edited`, 6);
    const fresh = await projectPptxToPpj(candidate.file, { sourceUri: `${type}-${suffix}-edited.pptx`, assetRootUri: "assets" });
    const chart = JSON.parse(new TextDecoder().decode(fresh.programJson)).pages[0].elements.find(e => e.type === "chart");
    assert.equal(chart.style.stacking, grouping);
    assert.deepEqual(chart.data.series[0].values, [6, null, 0, lastFirst]); assert.deepEqual(chart.data.series[1].values, [8, 2, 0, lastSecond]);
    const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file), changedParts = [];
    assert.deepEqual(Object.keys(newZip.files).sort(), Object.keys(oldZip.files).sort());
    for (const part of Object.keys(oldZip.files)) if (!oldZip.files[part].dir &&
      !Buffer.from(await oldZip.file(part).async("uint8array")).equals(Buffer.from(await newZip.file(part).async("uint8array")))) changedParts.push(part);
    assert.deepEqual(changedParts, ["ppt/slides/charts/chart1.xml"]); assert.deepEqual(source, original);
    nativeStacks.push({ type, grouping, authored: true, sourceNoop: true, sourceValueEdit: true, cumulativePixels: true, incompleteCategory: true, ...(percent ? { zeroTotalUndefined: true } : {}),
      reprojection: true, changedParts, sourceSha256: sha256(source), candidateSha256: sha256(candidate.file) });
  }
  const nativeCircular = [];
  for (const type of ["pie", "doughnut"]) {
    const circularProgram = structuredClone(pairBase);
    circularProgram.pages[0].elements = [
      { id: "behind", type: "shape", frame: { x: 60, y: 80, width: 400, height: 300 }, geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#FFEEDD" } } },
      { id: "circular", type: "chart", chartType: type, title: "Observed shares", frame: { x: 60, y: 80, width: 400, height: 300 },
        style: { startAngle: 0, ...(type === "doughnut" ? { holeSize: 60 } : {}) },
        data: { categories: ["Small", "Unknown", "Zero", "Large"], series: [{ id: "shares", name: "Shares", values: [1, null, 0, 9],
          pointStyles: [{ index: 0, fill: { type: "solid", color: "#CC2200" } }, { index: 3, fill: { type: "solid", color: "#0044CC" } }] }] } },
    ];
    async function checkCircular(receipt, name, swapped = false, angle = 0, hole = 60) {
      const painted = await savePaint(name, receipt), svg = painted.pages[0].svg;
      const native = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "chart").native;
      assert.equal(native.type, type === "pie" ? 3 : 5);
      assert.equal(native.firstSliceAngle, angle);
      if (type === "doughnut") assert.equal(native.doughnutHoleSize, hole);
      assert.deepEqual(native.series[0].values, swapped ? [9, 0, 0, 1] : [1, 0, 0, 9]);
      assert.deepEqual(native.series[0].missingValueIndexes, [1]);
      assert.match(svg, new RegExp(`data-officekit-chart="${type}"`));
      assert.match(svg, /data-officekit-missing-point="1"/);
      assert.doesNotMatch(svg, /data-officekit-point="1"|data-officekit-slice="[12]"/);
      assert.match(svg, /data-officekit-point="2" data-officekit-value="0" data-officekit-fraction="0"/);
      const first = svg.match(/data-officekit-point="0" data-officekit-value="[^"]+" data-officekit-fraction="([^"]+)" data-officekit-start-angle="([^"]+)"/);
      assert.ok(Math.abs(+first[1] - (swapped ? .9 : .1)) < 1e-12);
      assert.equal(+first[2], angle);
      assert.ok(!painted.diagnostics.some(d => ["preview.scene.paint.chart-semantics", "preview.scene.paint.failed"].includes(d.reason)));
      const sample = degrees => [Math.floor(260 + 80 * Math.sin(degrees * Math.PI / 180)), Math.floor(236 - 80 * Math.cos(degrees * Math.PI / 180))];
      for (const [[left, top], rgb] of [[sample(angle + 18), [204, 34, 0]], [sample(angle + 180), swapped ? [204, 34, 0] : [0, 68, 204]],
        ...(type === "doughnut" ? [[[260, 236], [255, 238, 221]]] : [])]) {
        const pixel = await sharp(Buffer.from(svg)).extract({ left, top, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
        assert.deepEqual([...pixel], rgb, `${name} pixel ${left},${top}`);
      }
    }
    const compileCircular = program => compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) }, { includePreviewScene: true });
    const authored = await compileCircular(circularProgram);
    await checkCircular(authored, `${type}-authored`);
    const swapped = structuredClone(circularProgram);
    swapped.pages[0].elements[1].data.series[0].values = [9, null, 0, 1];
    await checkCircular(await compileCircular(swapped), `${type}-swapped`, true);
    const rotated = structuredClone(circularProgram);
    rotated.pages[0].elements[1].style.startAngle = 90;
    if (type === "doughnut") rotated.pages[0].elements[1].style.holeSize = 40;
    await checkCircular(await compileCircular(rotated), `${type}-rotated`, false, 90, 40);
    const circularSource = await withoutAuthoredSnapshot(authored.file), original = circularSource.slice();
    const projectedCircular = await projectPptxToPpj(circularSource, { sourceUri: `${type}-source.pptx`, assetRootUri: "assets" });
    assert.equal(projectedCircular.sourceBound, true);
    const circularInput = { program: projectedCircular.programJson, source: circularSource, assets: projectedCircular.assets };
    const noopCircular = await compilePpjWorkspace(circularInput, { includePreviewScene: true });
    assert.deepEqual(noopCircular.file, circularSource);
    await checkCircular(noopCircular, `${type}-source-noop`);
    const editedCircular = JSON.parse(new TextDecoder().decode(projectedCircular.programJson));
    editedCircular.pages[0].elements.find(e => e.type === "chart").data.series[0].values = [9, null, 0, 1];
    const circularCandidate = await compilePpjWorkspace({ ...circularInput, program: Buffer.from(JSON.stringify(editedCircular)) }, { includePreviewScene: true });
    await checkCircular(circularCandidate, `${type}-source-edited`, true);
    const projectedAgain = await projectPptxToPpj(circularCandidate.file, { sourceUri: `${type}-candidate.pptx`, assetRootUri: "assets" });
    assert.deepEqual(JSON.parse(new TextDecoder().decode(projectedAgain.programJson)).pages[0].elements.find(e => e.type === "chart").data.series[0].values, [9, null, 0, 1]);
    const sourceZip = await JSZip.loadAsync(circularSource), candidateZip = await JSZip.loadAsync(circularCandidate.file);
    assert.deepEqual(Object.keys(sourceZip.files).sort(), Object.keys(candidateZip.files).sort());
    const changed = [];
    for (const name of Object.keys(sourceZip.files)) if (!sourceZip.files[name].dir &&
      !Buffer.from(await sourceZip.file(name).async("uint8array")).equals(Buffer.from(await candidateZip.file(name).async("uint8array")))) changed.push(name);
    assert.deepEqual(changed.sort(), ["ppt/slides/charts/chart1.xml"]);
    assert.ok(!Object.keys(sourceZip.files).some(name => name.endsWith(".xlsx")));
    assert.deepEqual(circularSource, original);
    const explosionProgram = structuredClone(circularProgram);
    const explosionSeries = explosionProgram.pages[0].elements[1].data.series[0];
    explosionSeries.values = [1, null, 0, 1]; explosionSeries.explosion = 100;
    explosionSeries.pointStyles[0].explosion = 25;
    async function checkExplosion(receipt, name, seriesValue, pointValue) {
      const painted = await savePaint(name, receipt), svg = painted.pages[0].svg;
      const native = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "chart").native.series[0];
      assert.equal(native.explosion, seriesValue);
      assert.equal(native.pointStyles[0].explosion, pointValue);
      assert.deepEqual(native.values, [1, 0, 0, 1]); assert.deepEqual(native.missingValueIndexes, [1]);
      assert.ok(!painted.diagnostics.some(d => ["preview.scene.paint.chart-semantics", "preview.scene.paint.failed"].includes(d.reason)));
      const effective = [pointValue ?? seriesValue ?? 0, seriesValue ?? 0];
      const thickness = type === "pie" ? 1 : .4;
      const radius = 102 / (1 + thickness * Math.max(...effective) / 100), inner = type === "pie" ? 0 : radius * .6;
      const dx = effective.map(e => (radius - inner) * e / 100);
      const points = [...svg.matchAll(/data-officekit-point="(\d+)"[^>]*data-officekit-explosion="([^"]+)" data-officekit-explosion-owner="([^"]+)" data-officekit-offset-x="([^"]+)" data-officekit-offset-y="([^"]+)"/g)];
      assert.deepEqual(points.map(m => +m[1]), [0, 2, 3]);
      assert.ok(Math.abs(+points[0][4] - dx[0]) < 1e-10); assert.ok(Math.abs(+points[2][4] + dx[1]) < 1e-10);
      assert.equal(points[0][3], pointValue !== undefined ? "point" : seriesValue !== undefined ? "series" : "default");
      assert.equal(+points[1][4], 0); assert.equal(+points[1][5], 0);
      assert.doesNotMatch(svg, /data-officekit-slice="[12]"|data-officekit-point="1"/);
      for (const [left, top, rgb] of [
        [Math.floor(260 + dx[0] + radius * .8), 236, [204, 34, 0]],
        [Math.floor(260 - dx[1] - radius * .8), 236, [0, 68, 204]],
        ...(type === "doughnut" ? [[Math.floor(260 + dx[0] + inner - 4), 236, [255, 238, 221]]] : dx[0] > 8 ? [[264, 236, [255, 238, 221]]] : []),
      ]) {
        const pixel = await sharp(Buffer.from(svg)).extract({ left, top, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
        assert.deepEqual([...pixel], rgb, `${name} exploded pixel ${left},${top}`);
      }
      assert.equal(painted.diagnostics.some(d => d.reason === "preview.scene.paint.chart-explosion-layout"), effective.some(e => e > 0));
    }
    const explosionAuthored = await compileCircular(explosionProgram);
    await checkExplosion(explosionAuthored, `${type}-explosion-authored`, 100, 25);
    const explosionSource = await withoutAuthoredSnapshot(explosionAuthored.file), explosionOriginal = explosionSource.slice();
    const explosionProjected = await projectPptxToPpj(explosionSource, { sourceUri: `${type}-explosion-source.pptx`, assetRootUri: "assets" });
    assert.equal(explosionProjected.sourceBound, true);
    const explosionInput = { program: explosionProjected.programJson, source: explosionSource, assets: explosionProjected.assets };
    const explosionNoop = await compilePpjWorkspace(explosionInput, { includePreviewScene: true });
    assert.deepEqual(explosionNoop.file, explosionSource);
    await checkExplosion(explosionNoop, `${type}-explosion-noop`, 100, 25);
    const explosionEdits = [];
    for (const [name, seriesValue, pointValue] of [["point-zero", 100, 0], ["point-delete", 100, undefined],
      ["series-zero", 0, 25], ["series-delete", undefined, 25], ["both-delete", undefined, undefined]]) {
      // Every edit is reconstructed from the ORIGINAL source projection.
      const edit = JSON.parse(new TextDecoder().decode(explosionProjected.programJson));
      const series = edit.pages[0].elements.find(e => e.type === "chart").data.series[0];
      if (seriesValue === undefined) delete series.explosion; else series.explosion = seriesValue;
      if (pointValue === undefined) delete series.pointStyles[0].explosion; else series.pointStyles[0].explosion = pointValue;
      const candidate = await compilePpjWorkspace({ ...explosionInput, program: Buffer.from(JSON.stringify(edit)) }, { includePreviewScene: true });
      await checkExplosion(candidate, `${type}-explosion-${name}`, seriesValue, pointValue);
      const reprojected = await projectPptxToPpj(candidate.file, { sourceUri: `${type}-${name}.pptx`, assetRootUri: "assets" });
      const projectedSeries = JSON.parse(new TextDecoder().decode(reprojected.programJson)).pages[0].elements.find(e => e.type === "chart").data.series[0];
      assert.equal(projectedSeries.explosion, seriesValue); assert.equal(projectedSeries.pointStyles[0].explosion, pointValue);
      assert.deepEqual(projectedSeries.values, [1, null, 0, 1]);
      const originalZip = await JSZip.loadAsync(explosionSource), editedZip = await JSZip.loadAsync(candidate.file);
      assert.deepEqual(Object.keys(editedZip.files).sort(), Object.keys(originalZip.files).sort());
      const changedParts = [];
      for (const member of Object.keys(originalZip.files)) if (!originalZip.files[member].dir &&
        !Buffer.from(await originalZip.file(member).async("uint8array")).equals(Buffer.from(await editedZip.file(member).async("uint8array")))) changedParts.push(member);
      assert.deepEqual(changedParts.sort(), ["ppt/slides/charts/chart1.xml"]);
      assert.deepEqual(explosionSource, explosionOriginal);
      explosionEdits.push({ name, sourceSha256: sha256(explosionSource), candidateSha256: sha256(candidate.file), changedParts, reprojection: true, pixels: true });
    }
    nativeCircular.push({ type, authoredRatioSwap: true, rotated: true, sourceNoop: true, sourceValueEdit: true, reprojection: true,
      ratioPixels: true, transparentHole: type === "doughnut", changedParts: changed,
      explosion: { authored: true, sourceNoop: true, layout: "radial-review; exact Office spacing unverified", edits: explosionEdits },
      sourceSha256: sha256(circularSource), candidateSha256: sha256(circularCandidate.file), workbook: "not present in literal-data fixture" });
  }
  const vectorProgram = structuredClone(pairBase);
  vectorProgram.pages[0].elements = [
    { id: "sun", type: "chart", chartType: "sunburst", frame: { x: 40, y: 80, width: 360, height: 350 }, style: { sunburst: { rootColors: ["#CC5500", "#114477"] } }, data: {
      categories: ["Root", "A", "B"], series: [{ id: "tree", name: "Tree", values: [10, 3, 7], parents: [null, "Root", "Root"] }] } },
    { id: "flow", type: "chart", chartType: "sankey", frame: { x: 450, y: 80, width: 450, height: 350 }, style: { sankey: { nodeColors: ["#CC5500", "#114477"] } }, data: {
      categories: ["A", "B", "C", "D"], series: [{ id: "flows", name: "Flow", values: [4, 6, 10], sources: ["A", "B", "C"], targets: ["C", "C", "D"] }] } },
  ];
  const vectorReceipt = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(vectorProgram)) }, { includePreviewScene: true });
  const vectorPaint = await savePaint("vectors", vectorReceipt), vectorView = createPpjSceneView(vectorReceipt);
  let paths = 0;
  function verifyPaths(node) {
    if (node.kind === "shape") for (const path of node.native.customPaths) {
      assert.ok(path.commands.some(c => c.command.case === "cubicBezierTo"));
      assert.ok(vectorPaint.pages[0].svg.includes(`d="${nativePathData(path, node.frame)}"`));
      assert.ok(node.semanticId === "sun" || node.semanticId === "flow");
      paths += 1;
    }
    node.children.forEach(verifyPaths);
  }
  vectorView.pages[0].nodes.forEach(verifyPaths);
  assert.ok(paths >= 5, `Expected actual generated sector and ribbon paths, got ${paths}`);
  assert.ok(!vectorPaint.diagnostics.some(d => ["preview.scene.paint.path", "preview.scene.paint.failed"].includes(d.reason)), JSON.stringify(vectorPaint.diagnostics));
  const gridInput = { id: "merged-grid", type: "table", frame: { x: 40, y: 100, width: 360, height: 120 },
    columns: [{ id: "narrow", width: 120 }, { id: "wide", width: 240 }], rows: [
      { id: "header", height: 40, cells: [{ id: "merged", text: "Merged header", columnSpan: 2, fill: { type: "solid", color: "#114477" } }] },
      { id: "values", height: 80, cells: [{ id: "zero", text: "0", fill: { type: "solid", color: "#00AA44" }, borders: { right: { color: "#0000FF", width: 2 } } },
        { id: "wide-value", text: "Wider", fill: { type: "solid", color: "#CC5500" } }] },
    ] };
  // A separate coordinate-endpoint case verifies the supported compile/painter
  // path without substituting for the required object-anchor cases below.
  const literalEdge = { id: "literal-edge", type: "connector", connectorType: "straight", frame: { x: 510, y: 120, width: 240, height: 200 },
    from: { x: 750, y: 320 }, to: { x: 510, y: 120 }, stroke: { color: "#114477", width: 2 }, endArrow: "triangle" };
  function assertLiteralEdge(receipt, painted) {
    const edge = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "connector");
    assert.deepEqual(edge.endpoints, { start: { x: 750, y: 320 }, end: { x: 510, y: 120 } });
    assert.ok(painted.pages[0].svg.includes('data-officekit-connector="straight" d="M 750 320 L 510 120"'));
    assert.ok(painted.pages[0].svg.includes('data-officekit-arrow="end" data-officekit-arrow-kind="triangle" transform="translate(510 120) rotate(-'));
  }
  const literalProgram = structuredClone(pairBase);
  literalProgram.pages[0].elements = [structuredClone(literalEdge)];
  const literalResult = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(literalProgram)) }, { includePreviewScene: true });
  assertLiteralEdge(literalResult, await savePaint("literal-connector", literalResult));
  async function assertMergedGrid(painted, x) {
    const svg = painted.pages[0].svg;
    assert.ok(svg.includes(`data-officekit-table-cell="0:0" data-officekit-row-span="1" data-officekit-column-span="2"><rect x="${x}" y="100" width="360" height="40"`));
    assert.match(svg, new RegExp(`data-officekit-table-cell="1:0"[^>]*><rect x="${x}" y="140" width="120" height="80"`));
    assert.match(svg, new RegExp(`data-officekit-table-cell="1:1"[^>]*><rect x="${x + 120}" y="140" width="240" height="80"`));
    assert.ok(svg.includes(`data-officekit-cell-border="right" x1="${x + 120}" y1="140" x2="${x + 120}" y2="220" stroke="#0000FF"`));
    assert.doesNotMatch(svg, /data-officekit-table-cell="0:1"/);
    assert.match(svg, />Merged header<\/tspan>/);
    assert.match(svg, />0<\/tspan>/);
    for (const [offset, expected] of [[40, [0, 170, 68]], [180, [204, 85, 0]]]) {
      const pixel = await sharp(Buffer.from(svg)).extract({ left: x + offset, top: 200, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
      assert.deepEqual([...pixel], expected);
    }
    assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.failed"), JSON.stringify(painted.diagnostics.filter(d => d.status === "unavailable")));
  }
  const topologyProgram = structuredClone(pairBase);
  topologyProgram.pages[0].elements = [structuredClone(gridInput),
    { ...structuredClone(primitive), id: "from-box", frame: { x: 750, y: 300, width: 60, height: 40 } },
    { ...structuredClone(primitive), id: "to-box", frame: { x: 450, y: 100, width: 60, height: 40 } },
    { id: "backward-edge", type: "connector", connectorType: "straight", frame: { x: 0, y: 0, width: 1, height: 1 },
      from: { element: "from-box", anchor: "left" }, to: { element: "to-box", anchor: "right" }, stroke: { color: "#114477", width: 2 }, endArrow: "triangle" },
  ];
  for (const shift of [0, 40]) {
    const program = structuredClone(topologyProgram);
    program.pages[0].elements[2].frame.y += shift;
    const result = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) }, { includePreviewScene: true });
    const painted = await savePaint(`topology-${shift}`, result), scene = createPpjSceneView(result);
    const edge = scene.pages[0].nodes.find(n => n.kind === "connector");
    // Independent table/source assertions must still execute when the known
    // authored-anchor defect fails. Keep every relation failure fatal at the
    // end; never replace object anchors with explicit points or bless fallback.
    try {
      assert.deepEqual(edge.endpoints, { start: { x: 750, y: 320 }, end: { x: 510, y: 120 + shift } });
      assert.ok(painted.pages[0].svg.includes(`d="M 750 320 L 510 ${120 + shift}"`));
      assert.ok(painted.pages[0].svg.includes(`data-officekit-arrow="end" data-officekit-arrow-kind="triangle" transform="translate(510 ${120 + shift}) rotate(-`));
      directedAnchorCases += 1;
    } catch (error) {
      relationFailures.push({ shift, actual: edge.endpoints, error });
    }
    await assertMergedGrid(painted, 40);
  }
  const json = JSON.parse(new TextDecoder().decode(sourceWorkspace.program));
  json.assets = JSON.parse(new TextDecoder().decode(authoredAssets.program)).assets;
  json.pages[0].elements.push({ id: "mark", type: "image", asset: json.assets[0].id,
    frame: { x: 80, y: 200, width: 100, height: 100 } });
  json.pages[0].elements.push(structuredClone(gridInput));
  json.pages[0].elements.push(structuredClone(literalEdge));
  json.pages[0].elements.push(structuredClone(nativeLineInput));
  const sourceAuthored = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(json)), assets: authoredAssets.assets });
  const source = await withoutAuthoredSnapshot(sourceAuthored.file);
  const beforeSource = source.slice();
  const projected = await projectPptxToPpj(source, { sourceUri: "source.pptx", assetRootUri: "assets" });
  assert.equal(projected.sourceBound, true);
  assert.equal(Object.hasOwn(projected, "previewScene"), false);
  const sourceInput = { program: projected.programJson, source, assets: projected.assets };
  const noop = await compilePpjWorkspace(sourceInput, { includePreviewScene: true });
  const noopView = createPpjSceneView(noop);
  assert.deepEqual(noopView.diagnostics, []);
  assert.ok(noopView.asset(noop.previewScene.assets[0].nativeId).data.byteLength);
  assert.deepEqual(noop.file, source);
  assert.equal(noop.previewScene.origin, 2);
  assert.ok(noop.previewScene.assets.length);
  assert.deepEqual((await fullWireCompile(sourceInput)).previewScene, noop.previewScene);
  const noopPaint = await savePaint("source-noop", noop);
  assertLiteralEdge(noop, noopPaint);
  await assertMergedGrid(noopPaint, 40);
  await assertNativeLine(noop, noopPaint);
  const edited = JSON.parse(new TextDecoder().decode(projected.programJson));
  edited.pages[0].elements[0].text.paragraphs[0].runs[0].text = "Actual candidate scene text";
  const editInput = { ...sourceInput, program: Buffer.from(JSON.stringify(edited)) };
  const candidate = await compilePpjWorkspace(editInput, { includePreviewScene: true });
  const candidateView = createPpjSceneView(candidate);
  const candidatePaint = await savePaint("source-edited", candidate);
  assert.match(candidatePaint.pages[0].svg, /Actual candidate scene text/);
  assert.equal(candidateView.pages[0].nodes[0].native.text, "Actual candidate scene text");
  assert.deepEqual(candidateView.diagnostics, []);
  assert.notEqual(candidate.outputSha256, projected.sourceSha256);
  assert.equal(candidate.previewScene.presentation.slides[0].elements[0].content.value.text, "Actual candidate scene text");
  assert.deepEqual((await fullWireCompile(editInput)).previewScene, candidate.previewScene);
  const tableEdit = JSON.parse(new TextDecoder().decode(projected.programJson));
  const target = tableEdit.pages[0].elements.find(e => e.type === "table");
  assert.ok(target, "Real source projection retains the bounded table instead of flattening it.");
  target.frame.x += 24;
  const tableCandidate = await compilePpjWorkspace({ ...sourceInput, program: Buffer.from(JSON.stringify(tableEdit)) }, { includePreviewScene: true });
  const tablePaint = await savePaint("source-table-moved", tableCandidate);
  assertLiteralEdge(tableCandidate, tablePaint);
  await assertMergedGrid(tablePaint, 64);
  await assertNativeLine(tableCandidate, tablePaint);
  assert.equal(createPpjSceneView(tableCandidate).pages[0].nodes.find(n => n.kind === "table").frame.x, 64);
  assert.match(tablePaint.pages[0].svg, /data-officekit-table-cell="0:0"[^>]*><rect x="64" y="100" width="360" height="40"/);
  const reprojected = await projectPptxToPpj(tableCandidate.file, { sourceUri: "candidate.pptx", assetRootUri: "assets" });
  assert.equal(JSON.parse(new TextDecoder().decode(reprojected.programJson)).pages[0].elements.find(e => e.type === "table").frame.x, 64);
  const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(tableCandidate.file);
  assert.deepEqual(Object.keys(newZip.files).filter(n => !newZip.files[n].dir).sort(), Object.keys(oldZip.files).filter(n => !oldZip.files[n].dir).sort());
  for (const name of Object.keys(oldZip.files)) if (!oldZip.files[name].dir && name !== "ppt/slides/slide1.xml")
    assert.deepEqual(await newZip.file(name).async("uint8array"), await oldZip.file(name).async("uint8array"), `Non-target part ${name}`);
  const chartEdit = JSON.parse(new TextDecoder().decode(projected.programJson));
  const chartTarget = chartEdit.pages[0].elements.find(e => e.type === "chart");
  assert.ok(chartTarget, "Source projection retains the bounded native line chart.");
  chartTarget.data.series[0].values[2] = 3;
  const chartCandidate = await compilePpjWorkspace({ ...sourceInput, program: Buffer.from(JSON.stringify(chartEdit)) }, { includePreviewScene: true });
  await assertNativeLine(chartCandidate, await savePaint("source-line-edited", chartCandidate), 3);
  const chartReprojected = await projectPptxToPpj(chartCandidate.file, { sourceUri: "line-candidate.pptx", assetRootUri: "assets" });
  assert.deepEqual(JSON.parse(new TextDecoder().decode(chartReprojected.programJson)).pages[0].elements.find(e => e.type === "chart").data.series[0].values, [2, null, 3, 4, 5]);
  assert.deepEqual(JSON.parse(new TextDecoder().decode(chartReprojected.programJson)).pages[0].elements.find(e => e.type === "chart").data.series[1].values, [5, 4, null, 0, 1]);
  const chartZip = await JSZip.loadAsync(chartCandidate.file);
  assert.deepEqual(Object.keys(chartZip.files).filter(n => !chartZip.files[n].dir).sort(), Object.keys(oldZip.files).filter(n => !oldZip.files[n].dir).sort());
  const chartChangedParts = [];
  for (const name of Object.keys(oldZip.files)) if (!oldZip.files[name].dir &&
    !Buffer.from(await oldZip.file(name).async("uint8array")).equals(Buffer.from(await chartZip.file(name).async("uint8array")))) chartChangedParts.push(name);
  // The fixture is a single literal-data ChartPart, without a workbook. Keep
  // this narrow golden delta: no slide, relationship, image or other part may
  // change. A workbook-linked source needs its own lifecycle regression.
  assert.ok(oldZip.file("ppt/slides/charts/chart1.xml"));
  assert.ok(!Object.keys(oldZip.files).some(name => name.endsWith(".xlsx")));
  assert.deepEqual(chartChangedParts.sort(), ["ppt/slides/charts/chart1.xml"]);
  assert.deepEqual(source, beforeSource);
  assert.ok(started.includes("office") && started.includes("ppj"));
  const report = { status: relationFailures.length ? "failed" : "passed", scope: "PPJ NativeAOT wire/view and internal SVG foundations; not production scene routing or complete paint coverage",
    nativeBars, nativeStacks, nativeCircular,
    relationFailures: relationFailures.map(({ shift, actual, error }) => ({ shift, actual, message: error.message })),
    internalPainting: { artifacts, pairedComponent: 1, customArcPath, generatedBezierPaths: paths, sourceTextEdit: true, directedAnchorCases, requiredDirectedAnchorCases: 2,
      explicitCoordinateConnector: true, sourceConnectorPreserved: true, mergedTablePixels: true, sourceTableMoveReprojection: true },
    nativeLine: { authored: true, sourceNoop: true, preservedAcrossTableEdit: true, sourceValueEditReprojection: true,
      multiSeriesMissing: true, markerAndGapPixels: true, unclippedMarkerOutlinePixels: true, changedParts: chartChangedParts,
      sourceSha256: sha256(source), candidateSha256: sha256(chartCandidate.file), workbook: "not present in this literal-data fixture" },
    officeProfile: "presentation request rejected as designed", profiles: descriptors.map(d => ({
    profile: d.profile, path: d.executablePath,
    sha256: d.manifest.files.find(file => file.path === d.manifest.profiles[d.profile].executable).sha256,
  })), authored: 2, sourceBound: ["no-op with assets/table", "text leaf edit with assets/table", "table frame edit and fresh reprojection", "literal line value edit and fresh reprojection"] };
  await writeFile(path.join(artifacts, "integration.json"), JSON.stringify(report, null, 2), { flag: "wx" });
  console.log(JSON.stringify(report, null, 2));
  if (relationFailures.length) throw new AggregateError(relationFailures.map(f => f.error), "Authored object-anchor regressions failed; independent table/source checks executed, not a passing integration.");
} finally {
  hook.deregister();
  delete globalThis[Symbol.for("officekit.preview.native.test")];
}
