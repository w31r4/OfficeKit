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
const fullWireClient = await startOfficeKitNativeClient({ packageJsonPath, profile: "ppj" });
try {
  const artifacts = await mkdtemp(path.join(os.tmpdir(), "officekit-native-scene-paint-"));
  console.log(`Native scene integration artifacts: ${artifacts}`);
  const relationFailures = [];
  let directedAnchorCases = 0;
  const { default: sharp } = await import("sharp");
  async function savePaint(name, receipt) {
    const painted = paintPpjSceneSvg(receipt);
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
  const fullWireCompile = async (workspace, { includePreviewScene = true, validationOnly = false } = {}) => {
    const request = create(CodecRequestSchema, {
    protocolVersion: 2, operation: 11, family: 2,
    presentationProgram: { programJson: workspace.program, includePreviewScene, validationOnly, includeNodeMap: true,
      assets: workspace.assets.map(asset => ({ id: asset.id, contentType: asset.mimeType, sha256: asset.sha256, data: asset.data })) },
    });
    const response = fromBinary(CodecResponseSchema,
      await fullWireClient.invoke(toBinary(CodecRequestSchema, request), workspace.source), { recursionLimit: 136 });
    assert.equal(response.ok, true, JSON.stringify(response.diagnostics));
    return { file: response.file, program: response.presentationProgram,
      ...(includePreviewScene ? { previewScene: readPpjPreviewScene(response.presentationProgram, response.file) } : {}) };
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
    geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } };
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
    assert.match(painted.pages[0].svg, /x="110" y="110" width="60" height="40" fill="#CC5500"/);
    assert.match(painted.pages[0].svg, />0<\/tspan>/);
    const pixel = await sharp(Buffer.from(painted.pages[0].svg)).extract({ left: 130, top: 130, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
    assert.deepEqual([...pixel], [204, 85, 0]);
    pair.push(leaves.map(n => ({ frame: n.frame, fill: n.native.fillRgb, text: n.native.text })));
  }
  assert.deepEqual(pair[0], pair[1]);
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
    assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.failed"));
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
  const sourceAuthored = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(json)), assets: authoredAssets.assets });
  const zip = await JSZip.loadAsync(sourceAuthored.file);
  for (const name of Object.keys(zip.files)) if (name.startsWith("officeKit/")) zip.remove(name);
  zip.file("_rels/.rels", (await zip.file("_rels/.rels").async("string")).replace(
    /<Relationship\b(?=[^>]*\bType="https:\/\/schemas\.officekit\.dev\/relationships\/presentation-program")[^>]*(?:\/>|>[\s\S]*?<\/Relationship>)/g, ""));
  zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml").async("string")).replace(
    /<Override\b(?=[^>]*\bPartName="\/officeKit\/)[^>]*(?:\/>|>[\s\S]*?<\/Override>)/g, ""));
  const source = await zip.generateAsync({ type: "uint8array" });
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
  assert.equal(createPpjSceneView(tableCandidate).pages[0].nodes.find(n => n.kind === "table").frame.x, 64);
  assert.match(tablePaint.pages[0].svg, /data-officekit-table-cell="0:0"[^>]*><rect x="64" y="100" width="360" height="40"/);
  const reprojected = await projectPptxToPpj(tableCandidate.file, { sourceUri: "candidate.pptx", assetRootUri: "assets" });
  assert.equal(JSON.parse(new TextDecoder().decode(reprojected.programJson)).pages[0].elements.find(e => e.type === "table").frame.x, 64);
  const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(tableCandidate.file);
  assert.deepEqual(Object.keys(newZip.files).filter(n => !newZip.files[n].dir).sort(), Object.keys(oldZip.files).filter(n => !oldZip.files[n].dir).sort());
  for (const name of Object.keys(oldZip.files)) if (!oldZip.files[name].dir && name !== "ppt/slides/slide1.xml")
    assert.deepEqual(await newZip.file(name).async("uint8array"), await oldZip.file(name).async("uint8array"), `Non-target part ${name}`);
  assert.deepEqual(source, beforeSource);
  assert.ok(started.includes("office") && started.includes("ppj"));
  console.log(JSON.stringify({ status: relationFailures.length ? "failed" : "passed", scope: "PPJ NativeAOT wire/view and internal SVG foundations; not production scene routing or complete paint coverage",
    relationFailures: relationFailures.map(({ shift, actual, error }) => ({ shift, actual, message: error.message })),
    internalPainting: { artifacts, pairedComponent: 1, generatedBezierPaths: paths, sourceTextEdit: true, directedAnchorCases, requiredDirectedAnchorCases: 2,
      explicitCoordinateConnector: true, sourceConnectorPreserved: true, mergedTablePixels: true, sourceTableMoveReprojection: true },
    officeProfile: "presentation request rejected as designed", profiles: descriptors.map(d => ({
    profile: d.profile, path: d.executablePath,
    sha256: d.manifest.files.find(file => file.path === d.manifest.profiles[d.profile].executable).sha256,
  })), authored: 2, sourceBound: ["no-op with assets/table", "text leaf edit with assets/table", "table frame edit and fresh reprojection"] }, null, 2));
  if (relationFailures.length) throw new AggregateError(relationFailures.map(f => f.error), "Authored object-anchor regressions failed; independent table/source checks executed, not a passing integration.");
} finally {
  fullWireClient.kill();
  hook.deregister();
  delete globalThis[Symbol.for("officekit.preview.native.test")];
}
