// Explicit integration against a checked-in-build-command output. This does not
// replace the installed package. Internal SVG foundations are exercised below;
// Selected production entry cases also execute real compilation/publication;
// complete painting and installed-package availability remain separate.
import assert from "node:assert/strict";
import path from "node:path";
import { registerHooks } from "node:module";
import JSZip from "jszip";
import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import { CodecRequestSchema, CodecResponseSchema, PresentationElementSchema } from "../src/generated/office_kit/artifact/v1/office_artifact_pb.js";
import { loadOfficeKitNativeDescriptor, startOfficeKitNativeClient } from "../src/codecs/office-kit-native-client.mjs";
import { readPpjPreviewScene } from "../src/ppj/preview-scene.mjs";
import { createPpjSceneView } from "../src/ppj/preview-scene-view.mjs";
import { paintPpjSceneSvg, nativePathData } from "../src/ppj/preview-scene-svg.mjs";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import { createHash } from "node:crypto";

assert.ok(process.argv[2], "Pass the directory produced by npm run build:office-kit -- --output <new-directory>.");
const evidenceFiles = ["../src/ppj/preview-scene-svg.mjs", "./ppj-preview-scene-native.mjs",
  "./ppj-preview-scene-svg.mjs", "./ppj-preview-render-assessment.mjs",
  "../src/ppj/svg-preview.mjs",
  "../src/ppj/preview-scene.mjs", "../src/ppj/preview-scene-view.mjs", "../src/ppj/preview-diagnostics.mjs", "../src/ppj/preview-output.mjs",
  "../src/ppj/preview-input-assessment.mjs", "../src/ppj/preview-factual-errors.mjs", "../src/ppj/capability-registry.json",
  "../src/ppj/preset-geometry-profiles.json",
  "../src/generated/office_kit/artifact/v1/office_artifact_pb.js",
  "./fixtures/presentation/preview-nested-repeat-equivalence.json", "./fixtures/presentation/preview-dataset-equivalence.json",
  "./fixtures/presentation/preview-style-grammar-equivalence.json"];
const javascriptIdentity = async () => Object.fromEntries(await Promise.all(evidenceFiles.map(async file =>
  [file, createHash("sha256").update(await readFile(new URL(file, import.meta.url))).digest("hex")])));
const javascriptAtStart = await javascriptIdentity();
const packageJsonPath = path.resolve(process.argv[2], "package.json");
const descriptors = await Promise.all(["office", "ppj"].map(profile => loadOfficeKitNativeDescriptor({ packageJsonPath, profile })));
const started = [];
let measuredTransport;
globalThis[Symbol.for("officekit.preview.native.test")] = async (options) => {
  const client = await startOfficeKitNativeClient({ ...options, packageJsonPath });
  started.push(options.profile);
  const invoke = client.invoke.bind(client);
  client.invoke = async (...args) => {
    const start = performance.now(), response = await invoke(...args);
    if (measuredTransport) measuredTransport.push({ requestBytes: args[0].byteLength,
      responseBytes: response.byteLength, invokeMs: performance.now() - start });
    return response;
  };
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
  const publicationCases = [];
  const publicationFailures = [];
  let directedAnchorCases = 0;
  let customArcPath = false;
  const { default: sharp } = await import("sharp");
  async function assertPublishedWarning(receipt, painted) {
    for (const [index, page] of receipt.pages.entries()) {
      assert.equal(page.reliability.status, painted.pages[index].reliability.status);
      const svg = await readFile(path.join(receipt.output.directory, page.file), "utf8");
      assert.match(svg, /INTERNAL SCENE PREVIEW/);
      const { data, info } = await sharp(await readFile(path.join(receipt.output.directory, page.png)))
        .removeAlpha().raw().toBuffer({ resolveWithObject: true });
      const offset = (info.width + info.width - 2) * info.channels;
      assert.deepEqual([...data.subarray(offset, offset + 3)],
        page.reliability.status === "failed" ? [153, 27, 27] : [146, 64, 14]);
    }
  }
  async function savePaint(name, receipt) {
    const painted = paintPpjSceneSvg(receipt);
    const nodes = new Map();
    function collect(assessment) {
      assert.ok(!nodes.has(assessment.scenePath), "each native node has one assessment");
      nodes.set(assessment.scenePath, assessment);
      assessment.children.forEach(collect);
    }
    for (const page of painted.pages) page.assessment.children.forEach(collect);
    assert.equal(nodes.size, receipt.previewScene.bindings.length);
    for (const binding of receipt.previewScene.bindings) {
      const assessment = nodes.get(binding.scenePath);
      assert.ok(assessment, binding.scenePath);
      assert.equal(assessment.pageId, binding.pageId);
      assert.equal(assessment.id, binding.semanticId || undefined);
      assert.equal(assessment.path, binding.programPath);
    }
    await writeFile(path.join(artifacts, `${name}.diagnostics.json`), JSON.stringify(painted.diagnostics, null, 2), { flag: "wx" });
    for (const [i, page] of painted.pages.entries()) {
      await writeFile(path.join(artifacts, `${name}-${i}.svg`), page.svg, { flag: "wx" });
      const png = await sharp(Buffer.from(page.svg)).png().toBuffer();
      await writeFile(path.join(artifacts, `${name}-${i}.png`), png, { flag: "wx" });
      assert.equal((await sharp(png).metadata()).width, painted.canvas.width);
    }
    return painted;
  }
  const { loadPpjWorkspace, compilePpjWorkspace, validatePpjWorkspace, sha256, writeExclusiveFile } = await import("../src/ppj/workspace.mjs");
  const { publishPpjPreview, previewInputEvidence } = await import("../src/ppj/preview-output.mjs");
  const { projectPptxToPpj } = await import("../src/ppj/native.mjs");
  const { renderPpjToSvg } = await import("../src/ppj/svg-preview.mjs");
  const { invokeOfficeKitLazy } = await import("../src/codecs/office-kit-runtime.mjs");
  const productionEntryCases = [];
  async function assertProductionEntry(name, input, compiled, painted) {
    const inputPath = path.join(artifacts, `entry-${name}.ppj`), outputDir = path.join(artifacts, `entry-${name}`);
    await writeFile(inputPath, input.program, { flag: "wx" });
    const sourcePath = input.source?.byteLength ? path.join(artifacts, `entry-${name}.source.pptx`) : undefined;
    if (sourcePath) await writeFile(sourcePath, input.source, { flag: "wx" });
    // Only loading is supplied in memory; compilation, scene validation,
    // mandatory combined assessment and SVG/PNG publication are the real path.
    const result = await renderPpjToSvg(inputPath, { outputDir,
      load: async () => ({ ...input, path: inputPath, sourcePath }),
    });
    assert.equal(result.renderer, "officekit-native-scene-svg");
    assert.equal(result.scene.sha256, compiled.previewScene.sha256);
    assert.equal(result.scene.candidateSha256, sha256(compiled.file));
    assert.ok(result.inputAssessment, "original-input checks are mandatory");
    assert.ok(!result.diagnostics.some(d => d.reason === "preview.scene.paint.integration-pending"));
    const persisted = JSON.parse(await readFile(path.join(outputDir, "render.json"), "utf8"));
    assert.deepEqual(result.receipt, persisted);
    assert.deepEqual(result.assessment, persisted.assessment);
    assert.doesNotThrow(() => JSON.stringify(result));
    for (const [i, page] of result.pages.entries()) {
      const actual = await readFile(path.join(outputDir, persisted.pages[i].png));
      const region = { left: 0, top: 24, width: result.canvas.width, height: result.canvas.height - 24 };
      assert.deepEqual(await sharp(actual).extract(region).raw().toBuffer(),
        await sharp(Buffer.from(painted.pages[i].svg)).extract(region).raw().toBuffer(),
        "public entry must consume the same proven native geometry/style, not canonical heuristics");
      const pixel = await sharp(actual).extract({ left: 1, top: 1, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
      assert.deepEqual([...pixel], page.reliability.status === "failed" ? [153, 27, 27] : [146, 64, 14]);
    }
    assert.deepEqual(await readFile(inputPath), Buffer.from(input.program));
    if (sourcePath) assert.deepEqual(await readFile(sourcePath), Buffer.from(input.source));
    productionEntryCases.push({ name, scene: result.scene, reliability: result.reliability.status,
      completeContentRasterEqual: true, mandatoryAssessment: true, returnedPersistedEqual: true, inputsPreserved: true });
    return result;
  }
  async function withoutAuthoredSnapshot(file) {
    const archive = await JSZip.loadAsync(file);
    for (const name of Object.keys(archive.files)) if (name.startsWith("officeKit/")) archive.remove(name);
    archive.file("_rels/.rels", (await archive.file("_rels/.rels").async("string")).replace(
      /<Relationship\b(?=[^>]*\bType="https:\/\/schemas\.officekit\.dev\/relationships\/presentation-program")[^>]*(?:\/>|>[\s\S]*?<\/Relationship>)/g, ""));
    archive.file("[Content_Types].xml", (await archive.file("[Content_Types].xml").async("string")).replace(
      /<Override\b(?=[^>]*\bPartName="\/officeKit\/)[^>]*(?:\/>|>[\s\S]*?<\/Override>)/g, ""));
    return archive.generateAsync({ type: "uint8array" });
  }
  const xmlAttributes = tag => Object.fromEntries([...tag.matchAll(/([\w:]+)="([^"]*)"/g)].map(m => [m[1], m[2]]));
  // Controlled SDK fixtures may reorder attributes and add the same a
  // namespace at the root. Preserve all other names, values and content.
  const orderedXml = xml => xml.replace(/<p:sld\b[^>]*>/, tag => tag.replace(' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"', ""))
    .replace(/<([\w:.-]+)(\s[^<>]*?)?(\/?)>/g, (tag, name, _, close) =>
    `<${name} ${JSON.stringify(Object.entries(xmlAttributes(tag)).sort(([a],[b]) => a.localeCompare(b)))}${close}>`);
  const fullWireCompile = async (workspace, { includePreviewScene = true, validationOnly = false, measureNativeMemory = false } = {}) => {
    const request = create(CodecRequestSchema, {
    protocolVersion: 2, operation: 11, family: 2,
    presentationProgram: { programJson: workspace.program, includePreviewScene, validationOnly, includeNodeMap: true,
      assets: workspace.assets.map(asset => ({ id: asset.id, contentType: asset.mimeType, sha256: asset.sha256, data: asset.data })) },
    });
    // Independent comparisons can be separated by a long raster/edit suite.
    // Own the client per comparison; never reuse a handle after idle retirement
    // or change production idle policy to keep this test's handle alive.
    const startupStart = performance.now();
    const client = await startOfficeKitNativeClient({ packageJsonPath, profile: "ppj" });
    const startupMs = performance.now() - startupStart;
    try {
      const memorySnapshot = async () => {
        if (process.platform !== "linux") return { status: "unavailable", reason: "Linux procfs measurement only" };
        const pid = client.child.pid;
        assert.ok(Number.isSafeInteger(pid) && pid > 0);
        const status = await readFile(`/proc/${pid}/status`, "utf8");
        const field = name => {
          const match = status.match(new RegExp(`^${name}:\\s+(\\d+)\\s+kB$`, "m"));
          assert.ok(match, `native ${name} measurement must exist`);
          return Number(match[1]) * 1024;
        };
        return { status: "measured", pid, rssBytes: field("VmRSS"), lifetimeHighWaterBytes: field("VmHWM") };
      };
      const before = measureNativeMemory ? await memorySnapshot() : undefined;
      const requestBytes = toBinary(CodecRequestSchema, request), invokeStart = performance.now();
      const responseBytes = await client.invoke(requestBytes, workspace.source);
      const invokeMs = performance.now() - invokeStart;
      const after = measureNativeMemory ? await memorySnapshot() : undefined;
      if (before?.status === "measured") {
        assert.equal(before.pid, after.pid);
        assert.ok(after.lifetimeHighWaterBytes >= before.lifetimeHighWaterBytes);
      }
      const response = fromBinary(CodecResponseSchema, responseBytes, { recursionLimit: 136 });
      assert.equal(response.ok, true, JSON.stringify(response.diagnostics));
      return { file: response.file, program: response.presentationProgram,
        ...(measureNativeMemory ? { nativeMemory: { before, after, startupMs, invokeMs,
          requestBytes: requestBytes.byteLength, responseBytes: responseBytes.byteLength,
          scope: "Fresh process: handshake plus one compile; OS resident high-water, not managed live-object retention. includeNodeMap=true." } } : {}),
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
    publicationCases.push({ name: fixture.includes("minimum") ? "minimum" : "canonical", input: workspace, compiled: ppj, painted });
    assert.match(painted.pages[0].svg, /data-officekit-native-id=/);
    const publicationDir = path.join(artifacts, fixture.includes("minimum") ? "published-minimum" : "published-canonical");
    const publication = await publishPpjPreview(painted, previewInputEvidence(workspace, ppj), { outputDir: publicationDir });
    const persisted = JSON.parse(await readFile(path.join(publicationDir, "render.json"), "utf8"));
    assert.deepEqual(persisted, JSON.parse(JSON.stringify(publication.receipt)));
    assert.deepEqual(persisted.diagnostics, JSON.parse(JSON.stringify(painted.diagnostics)));
    assert.deepEqual(persisted.assessment, JSON.parse(JSON.stringify(painted.assessment)));
    assert.deepEqual(persisted.pages.map(p => p.id), painted.pages.map(p => p.id));
    assert.deepEqual(persisted.scene, painted.sceneEvidence);
    assert.equal(persisted.scene.origin, "authored-lowering");
    assert.equal(persisted.scene.candidateSha256, sha256(ppj.file));
    assert.equal(persisted.scene.sha256, ppj.previewScene.sha256);
    await assertPublishedWarning(persisted, painted);
    assert.throws(() => previewInputEvidence(workspace, { ...ppj, file: Buffer.from("wrong candidate") }),
      error => error.code === "preview.scene.candidate-mismatch");
    assert.throws(() => previewInputEvidence(workspace, { ...ppj, previewScene: { ...ppj.previewScene, sha256: "0".repeat(64) } }),
      error => error.code === "preview.scene.digest-mismatch");
    if (ppj.previewScene.assets.length) assert.throws(() => previewInputEvidence(workspace, { ...ppj, assets: [] }),
      error => error.code === "preview.scene.asset-mismatch");
    assert.ok(persisted.implementation.sources.some(s => s.file === "preview-scene-svg.mjs"));
    for (const artifact of persisted.artifacts) assert.equal(sha256(await readFile(path.join(publicationDir, artifact.file))), artifact.sha256);
    await assert.rejects(publishPpjPreview(painted, previewInputEvidence(workspace, ppj), { outputDir: publicationDir }),
      error => error.code === "preview.output.exists");
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
  // Real rich-text decoration/baseline mapping, independent of font-specific
  // glyph contours: compare identical glyphs and inspect their raster bounds.
  const formatProgram = structuredClone(pairBase);
  formatProgram.pages[0].elements = [{ id: "formatted", type: "text", frame: { x: 100, y: 100, width: 300, height: 100 },
    text: { paragraphs: [{ runs: [{ text: "HHHH", style: { size: 40, color: "#000000", baseline: 0, letterSpacing: 0, underline: "none", strike: "noStrike" } }] }] } }];
  const compileFormat = program => compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) }, { includePreviewScene: true });
  const gradientCases = [];
  const gradientStops = [
    { offset: 0, color: "#FF0000" }, { offset: .25, color: "#FF0000" },
    { offset: .25, color: "#00FF00", opacity: 0 }, { offset: .75, color: "#00FF00", opacity: 0 },
    { offset: .75, color: "#0000FF" }, { offset: 1, color: "#0000FF" },
  ];
  async function gradientPixels(name, receipt, shapeAngle, tableAngle = shapeAngle) {
    const painted = await savePaint(`gradient-${name}`, receipt), svg = painted.pages[0].svg;
    assert.equal(svg.match(/<linearGradient/g)?.length, 2);
    assert.equal(new Set([...svg.matchAll(/<linearGradient id="([^"]+)"/g)].map(m => m[1])).size, 2);
    const view = createPpjSceneView(receipt), shape = view.pages[0].nodes.find(n => n.kind === "shape").native;
    const table = view.pages[0].nodes.find(n => n.kind === "table").native;
    for (const [g, angle] of [[shape.gradientFill, shapeAngle], [table.rows[0].cells[0].fill.kind.value, tableAngle]]) {
      assert.equal(g.angle60000, angle * 60000);
      assert.deepEqual(g.stops.map(s => s.positionThousandthPercent), [0, 25000, 25000, 75000, 75000, 100000]);
      assert.equal(g.stops[2].opacityThousandthPercent, 0);
    }
    const raster = await sharp(Buffer.from(svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    for (const [x, angle] of [[100, shapeAngle], [400, tableAngle]]) for (const [fraction, rgb] of [[.1, [255, 0, 0]], [.5, [255, 255, 255]], [.9, [0, 0, 255]]]) {
      const px = x + (angle === 90 ? 100 : 200 * fraction), py = 100 + (angle === 90 ? 100 * fraction : 50);
      const offset = (py * raster.info.width + px) * raster.info.channels;
      assert.deepEqual([...raster.data.subarray(offset, offset + 3)], rgb);
    }
    assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.gradient" || d.reason === "preview.scene.paint.failed"));
    gradientCases.push({ name, shapeAngle, tableAngle, plateauAndTransparentPixels: true, candidateSha256: sha256(receipt.file) });
  }
  for (const angle of [0, 45, 90]) {
    const program = structuredClone(pairBase), fill = { type: "gradient", kind: "linear", angle, stops: structuredClone(gradientStops) };
    program.pages[0].elements = [
      { id: "gradient-shape", type: "shape", frame: { x: 100, y: 100, width: 200, height: 100 }, geometry: { kind: "preset", preset: "rect" }, style: { fill } },
      { id: "gradient-table", type: "table", frame: { x: 400, y: 100, width: 200, height: 100 }, columns: [{ id: "col", width: 200 }],
        rows: [{ id: "row", height: 100, cells: [{ id: "cell", text: "", fill: structuredClone(fill) }] }] },
    ];
    const originalProgram = JSON.stringify(program), authored = await compileFormat(program);
    await gradientPixels(`${angle}-authored`, authored, angle);
    const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    const projected = await projectPptxToPpj(source, { sourceUri: "gradient.pptx", assetRootUri: "assets" });
    const input = { program: projected.programJson, source, assets: projected.assets };
    const noop = await compilePpjWorkspace(input, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    await gradientPixels(`${angle}-source`, noop, angle);
    if (angle === 90) {
      const request = JSON.parse(Buffer.from(projected.programJson).toString("utf8"));
      request.pages[0].elements[0].style.fill.angle = 0;
      const candidate = await compilePpjWorkspace({ ...input, program: Buffer.from(JSON.stringify(request)) }, { includePreviewScene: true });
      await gradientPixels("source-edit", candidate, 0, 90);
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: "gradient-edited.pptx", assetRootUri: "assets" });
      assert.equal(JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0].elements[0].style.fill.angle, 0);
      const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file), changed = [];
      assert.deepEqual(Object.keys(oldZip.files).sort(), Object.keys(newZip.files).sort());
      for (const name of Object.keys(oldZip.files)) if (!oldZip.files[name].dir &&
        !Buffer.from(await oldZip.file(name).async("uint8array")).equals(Buffer.from(await newZip.file(name).async("uint8array")))) changed.push(name);
      assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
    }
    assert.equal(sha256(source), sourceHash);
    assert.equal(JSON.stringify(program), originalProgram);
  }
  const backgroundGradientCases = [], backgroundGradientFailures = [];
  const backgroundImageCases = [], backgroundImageFailures = [], backgroundImageRejections = [];
  const shapeImageCases = [], shapeImageFailures = [];
  const polygonPresetCases = [], polygonPresetFailures = [];
  const radialGradientCases = [], radialGradientFailures = [], radialGeometryCases = [];
  const brightnessCases = [], brightnessFailures = [];
  async function backgroundGradientPixels(name, receipt, angle, alpha = 0, removed = false) {
    const painted = await savePaint(`background-gradient-${name}`, receipt), svg = painted.pages[0].svg;
    const background = createPpjSceneView(receipt).pages[0].native.background;
    if (removed) {
      assert.equal(background?.gradientFill, undefined);
      assert.ok(!svg.includes("<linearGradient"));
    } else {
      assert.equal(background.gradientFill.kind, 1);
      assert.equal(background.gradientFill.angle60000, angle * 60000);
      assert.deepEqual(background.gradientFill.stops.map(s => s.positionThousandthPercent), [0, 25000, 25000, 75000, 75000, 100000]);
      assert.equal(background.gradientFill.stops[2].opacityThousandthPercent, alpha * 100000);
      assert.equal(background.gradientFill.stops[3].opacityThousandthPercent, alpha * 100000);
      assert.equal(svg.match(/<linearGradient/g)?.length, 1);
      assert.match(svg, /data-officekit-background="gradient"/);
    }
    const raster = await sharp(Buffer.from(svg)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
    const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * 4, (y * raster.info.width + x) * 4 + 4)];
    for (const [fraction, color] of [[.1, [255, 0, 0]], [.5, [0, 255, 0]], [.9, [0, 0, 255]]]) {
      const x = Math.floor(raster.info.width * (angle === 90 ? .5 : fraction));
      const y = Math.floor(raster.info.height * (angle === 0 ? .5 : fraction));
      const rgba = pixel(x, y);
      if (removed) assert.deepEqual(rgba, [255, 255, 255, 255]);
      else if (fraction === .5) {
        assert.equal(rgba[3], Math.round(alpha * 255), "transparent background alpha is not flattened onto white");
        if (alpha > 0) assert.deepEqual(rgba.slice(0, 3), color);
      } else assert.deepEqual(rgba, [...color, 255]);
    }
    assert.deepEqual(pixel(120, 120), [204, 85, 0, 255], "foreground stays above the gradient");
    assert.ok(!painted.diagnostics.some(d => d.scenePath.endsWith(".background.gradientFill") || d.reason === "preview.scene.paint.background"));
    backgroundGradientCases.push({ name, angle, alpha, removed, nativeAndRgbaPixels: true, foregroundPreserved: true,
      candidateSha256: sha256(receipt.file), sceneSha256: receipt.previewScene.sha256 });
    return painted;
  }
  for (const angle of [0, 45, 90]) try {
    const program = structuredClone(pairBase);
    program.pages[0].background = { type: "gradient", kind: "linear", angle, stops: structuredClone(gradientStops) };
    program.pages[0].elements = [{ id: "background-control", type: "shape", frame: { x: 100, y: 100, width: 60, height: 40 },
      geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } }];
    const input = { ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) }, inputHash = sha256(input.program);
    const authored = await compilePpjWorkspace(input, { includePreviewScene: true });
    const originalPaint = await backgroundGradientPixels(`${angle}-authored`, authored, angle);
    await assertProductionEntry(`background-${angle}-authored`, input, authored, originalPaint);
    const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    await writeFile(path.join(artifacts, `background-${angle}-source.pptx`), source, { flag: "wx" });
    const projected = await projectPptxToPpj(source, { sourceUri: `background-${angle}.pptx`, assetRootUri: "assets" });
    const bound = { program: projected.programJson, source, assets: projected.assets };
    const noop = await compilePpjWorkspace(bound, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    const sourcePaint = await backgroundGradientPixels(`${angle}-source`, noop, angle);
    await assertProductionEntry(`background-${angle}-source`, bound, noop, sourcePaint);
    if (angle === 90) for (const edit of ["angle", "alpha", "delete"]) {
      // Every request starts from the same original projection, never a prior
      // candidate. Deletion must remove native/PPJ state, not merely look white.
      const request = JSON.parse(Buffer.from(projected.programJson).toString("utf8"));
      if (edit === "delete") delete request.pages[0].background;
      if (edit === "angle") request.pages[0].background.angle = 0;
      if (edit === "alpha") for (const i of [2, 3]) request.pages[0].background.stops[i].opacity = .5;
      const edited = { ...bound, program: Buffer.from(JSON.stringify(request)) }, requestHash = sha256(edited.program);
      const candidate = await compilePpjWorkspace(edited, { includePreviewScene: true });
      const painted = await backgroundGradientPixels(`source-${edit}`, candidate, edit === "angle" ? 0 : 90, edit === "alpha" ? .5 : 0, edit === "delete");
      await assertProductionEntry(`background-source-${edit}`, edited, candidate, painted);
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: `background-${edit}.pptx`, assetRootUri: "assets" });
      const observed = JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0].background;
      if (edit === "delete") assert.equal(observed, undefined);
      if (edit === "angle") assert.equal(observed.angle, 0);
      if (edit === "alpha") for (const i of [2, 3]) assert.equal(observed.stops[i].opacity, .5);
      const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file), changed = [];
      assert.deepEqual(Object.keys(newZip.files).sort(), Object.keys(oldZip.files).sort());
      for (const name of Object.keys(oldZip.files)) if (!oldZip.files[name].dir &&
        !Buffer.from(await oldZip.file(name).async("uint8array")).equals(Buffer.from(await newZip.file(name).async("uint8array")))) changed.push(name);
      assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
      if (edit === "delete") assert.ok(!/<p:bg[ >]/.test(await newZip.file("ppt/slides/slide1.xml").async("string")));
      assert.equal(sha256(edited.program), requestHash);
      Object.assign(backgroundGradientCases.at(-1), { reprojection: true, changedParts: changed,
        sourceSha256: sourceHash, requestSha256: requestHash });
    }
    assert.equal(sha256(source), sourceHash);
    assert.equal(sha256(input.program), inputHash);
  } catch (error) {
    backgroundGradientFailures.push({ angle, code: error.code, message: error.message });
    console.error(`Background gradient ${angle} failed: ${error.message}`);
  }
  // One original source contains all three consumers. Every edit starts from
  // that original projection and must preserve the other two gradients.
  try {
    const radialFill = { type: "gradient", kind: "radial", stops: structuredClone(gradientStops) };
    const program = structuredClone(pairBase);
    program.pages[0].background = structuredClone(radialFill);
    program.pages[0].elements = [
      { id: "radial-shape", type: "shape", frame: { x: 200, y: 140, width: 200, height: 100 },
        geometry: { kind: "preset", preset: "rect" }, style: { fill: structuredClone(radialFill), stroke: { color: "#FF00FF", width: 2 } },
        text: { paragraphs: [{ runs: [{ text: "HHHH", style: { size: 12, color: "#000000" } }] }] } },
      { id: "radial-table", type: "table", frame: { x: 600, y: 140, width: 200, height: 100 }, columns: [{ id: "col", width: 200 }],
        rows: [{ id: "row", height: 100, cells: [{ id: "cell", text: "", fill: structuredClone(radialFill) }] }] },
      { id: "radial-control", type: "shape", frame: { x: 40, y: 350, width: 20, height: 20 },
        geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } },
    ];
    const input = { ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) }, inputHash = sha256(input.program);
    const authored = await compilePpjWorkspace(input, { includePreviewScene: true });
    const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    await writeFile(path.join(artifacts, "radial-source.pptx"), source, { flag: "wx" });
    const projection = await projectPptxToPpj(source, { sourceUri: "radial-source.pptx", assetRootUri: "assets" });
    const projectionHash = sha256(projection.programJson), originalZip = await JSZip.loadAsync(source);
    const bound = { source, assets: projection.assets, program: projection.programJson };
    const noop = await compilePpjWorkspace(bound, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    const fillAt = (p, owner) => owner === "background" ? p.pages[0].background : owner === "shape"
      ? p.pages[0].elements[0].style.fill : p.pages[0].elements[1].rows[0].cells[0].fill;
    async function radialPixels(name, receipt, owner, edit) {
      const painted = await savePaint(`radial-${name}`, receipt), view = createPpjSceneView(receipt), page = view.pages[0];
      const shape = page.nodes[0].native, tableFill = page.nodes.find(n => n.kind === "table").native.rows[0].cells[0].fill;
      const fills = { shape: shape.gradientFill, table: tableFill?.kind.case === "gradientFill" ? tableFill.kind.value : undefined,
        background: page.native.background?.gradientFill };
      for (const [consumer, fill] of Object.entries(fills)) {
        if (consumer === owner && edit === "delete") { assert.equal(fill, undefined); continue; }
        const linear = consumer === owner && edit === "linear";
        assert.equal(fill.kind, linear ? 1 : 2);
        assert.equal(fill.angle60000, linear ? 0 : undefined);
        assert.deepEqual(fill.stops.map(s => s.positionThousandthPercent), [0, 25000, 25000, 75000, 75000, 100000]);
        assert.deepEqual(fill.stops.map(s => s.colorRgb), [0, 1].map(() => consumer === owner && edit === "color" ? "FFFF00" : "FF0000").concat(["00FF00", "00FF00", "0000FF", "0000FF"]));
        for (const index of [2, 3]) assert.equal(fill.stops[index].opacityThousandthPercent, consumer === owner && edit === "alpha" ? 50000 : 0);
      }
      const svg = painted.pages[0].svg;
      assert.equal((svg.match(/<radialGradient/g) || []).length, edit === "delete" || edit === "linear" ? 2 : 3);
      assert.equal((svg.match(/<linearGradient/g) || []).length, edit === "linear" ? 1 : 0);
      const raster = await sharp(Buffer.from(svg)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
      const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * 4, (y * raster.info.width + x) * 4 + 4)];
      const red = [255, 0, 0, 255], blue = [0, 0, 255, 255], clear = [0, 0, 0, 0], green = [0, 255, 0, 128], white = [255, 255, 255, 255];
      const underlay = owner === "background" && edit === "delete" ? white : owner === "background" && edit === "alpha" ? green : clear;
      for (const [consumer, cx, cy] of [["shape", 300, 190], ["table", 700, 190], ["background", 480, 270]]) {
        const changed = consumer === owner;
        let center = changed && edit === "color" ? [255, 255, 0, 255] : red;
        let middle = consumer === "background" ? clear : underlay;
        let edge = blue;
        if (changed && edit === "alpha") middle = green;
        if (changed && edit === "delete") center = middle = edge = consumer === "background" ? white : clear;
        if (changed && edit === "linear") { center = middle; edge = red; }
        // Equal physical-distance X/Y probes distinguish the circumscribed
        // circle from an inscribed ellipse on these non-square frames.
        assert.deepEqual(pixel(cx, cy), center, `${name}/${consumer}: center stop`);
        assert.deepEqual(pixel(cx, consumer === "background" ? 50 : 145), middle, `${name}/${consumer}: transparent/alpha ring`);
        if (consumer !== "background") for (const x of [cx - 45, cx + 45]) {
          // The transparent right-hand table sample reveals the blue final
          // plateau of the edited linear background: 745/960 > 0.75.
          const expected = owner === "background" && edit === "linear" && consumer === "table" && x === 745 ? blue : middle;
          assert.deepEqual(pixel(x, cy), expected, `${name}/${consumer}: equal physical radius at x=${x}`);
        }
        assert.deepEqual(pixel(consumer === "background" ? 30 : cx - 95, consumer === "background" ? 30 : cy), edge, `${name}/${consumer}: outer stop`);
      }
      assert.deepEqual(pixel(50, 360), [204, 85, 0, 255]);
      assert.deepEqual(pixel(300, 140), [255, 0, 255, 255]);
      assert.match(svg, /HHHH/);
      assert.ok(!painted.diagnostics.some(d => ["preview.scene.paint.gradient", "preview.scene.paint.background"].includes(d.reason)));
      return painted;
    }
    for (const [name, receipt, workspace] of [["authored", authored, input], ["source", noop, bound]]) {
      const painted = await radialPixels(name, receipt);
      await assertProductionEntry(`radial-${name}`, workspace, receipt, painted);
      radialGradientCases.push({ name, allThreeConsumers: true, nativeAndRgbaPixels: true, sourceNoop: name === "source",
        candidateSha256: sha256(receipt.file), sceneSha256: receipt.previewScene.sha256, sourceSha256: sourceHash });
    }
    for (const owner of ["shape", "table", "background"]) for (const edit of ["color", "alpha", "linear", "delete"]) {
      const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8")), fill = fillAt(request, owner);
      if (edit === "color") for (const index of [0, 1]) fill.stops[index].color = "#FFFF00";
      if (edit === "alpha") for (const index of [2, 3]) fill.stops[index].opacity = .5;
      if (edit === "linear") { fill.kind = "linear"; fill.angle = 0; }
      if (edit === "delete") {
        if (owner === "background") delete request.pages[0].background;
        else if (owner === "shape") delete request.pages[0].elements[0].style.fill;
        else delete request.pages[0].elements[1].rows[0].cells[0].fill;
      }
      const name = `${owner}-${edit}`, edited = { ...bound, program: Buffer.from(JSON.stringify(request)) };
      const requestHash = sha256(edited.program), requestFile = `radial-${name}.ppj`, evidence = { name, stage: "compile", sourceSha256: sourceHash, requestSha256: requestHash, requestFile };
      await writeFile(path.join(artifacts, requestFile), edited.program, { flag: "wx" });
      try {
        const candidate = await compilePpjWorkspace(edited, { includePreviewScene: true });
        const candidateFile = `radial-${name}.pptx`;
        await writeFile(path.join(artifacts, candidateFile), candidate.file, { flag: "wx" });
        Object.assign(evidence, { stage: "paint", candidateFile, candidateSha256: sha256(candidate.file) });
        const painted = await radialPixels(name, candidate, owner, edit);
        await assertProductionEntry(`radial-${name}`, edited, candidate, painted);
        evidence.stage = "reprojection";
        const fresh = await projectPptxToPpj(candidate.file, { sourceUri: candidateFile, assetRootUri: "assets" });
        const observed = JSON.parse(Buffer.from(fresh.programJson).toString("utf8"));
        for (const consumer of ["shape", "table", "background"]) assert.deepEqual(fillAt(observed, consumer), fillAt(request, consumer), `${name}: fresh ${consumer} fill`);
        const candidateZip = await JSZip.loadAsync(candidate.file), changed = [];
        assert.deepEqual(Object.keys(candidateZip.files).sort(), Object.keys(originalZip.files).sort());
        for (const part of Object.keys(originalZip.files)) if (!originalZip.files[part].dir &&
          !Buffer.from(await originalZip.file(part).async("uint8array")).equals(Buffer.from(await candidateZip.file(part).async("uint8array")))) changed.push(part);
        assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
        radialGradientCases.push({ ...evidence, stage: "verified", nativeAndRgbaPixels: true, allThreeConsumers: true, reprojection: true, changedParts: changed });
      } catch (error) {
        radialGradientFailures.push({ ...evidence, code: error.code, message: error.message });
        console.error(`Radial gradient ${name} failed at ${evidence.stage}: ${error.message}`);
      }
      assert.equal(sha256(source), sourceHash); assert.equal(sha256(edited.program), requestHash);
    }
    const geometryFixtures = [
      { name: "partial-rect", width: 100, height: 100, commands: [
        { op: "moveTo", x: 0, y: 0 }, { op: "lineTo", x: 50, y: 0 },
        { op: "lineTo", x: 50, y: 100 }, { op: "lineTo", x: 0, y: 100 }, { op: "close" }] },
      { name: "quadratic", width: 200, height: 50, commands: [
        { op: "moveTo", x: 0, y: 0 }, { op: "quadraticTo", x1: 50, y1: 100, x: 100, y: 0 }, { op: "close" }] },
      { name: "cubic", width: 200, height: 75, commands: [
        { op: "moveTo", x: 0, y: 0 }, { op: "cubicTo", x1: 0, y1: 100, x2: 100, y2: 100, x: 100, y: 0 }, { op: "close" }] },
    ];
    for (const geometry of geometryFixtures) {
      try {
        const geometryProgram = structuredClone(program);
        geometryProgram.pages[0].elements[0].geometry = { kind: "custom", viewBox: { x: 0, y: 0, width: 100, height: 100 },
          paths: [{ fill: true, stroke: true, commands: geometry.commands }] };
        const geometryInput = { ...sourceWorkspace, program: Buffer.from(JSON.stringify(geometryProgram)) };
        const geometryInputHash = sha256(geometryInput.program);
        const author = await compilePpjWorkspace(geometryInput, { includePreviewScene: true });
        const sourceBytes = await withoutAuthoredSnapshot(author.file), sourceDigest = sha256(sourceBytes);
        const sourceFile = `radial-geometry-${geometry.name}-source.pptx`;
        await writeFile(path.join(artifacts, sourceFile), sourceBytes, { flag: "wx" });
        const projected = await projectPptxToPpj(sourceBytes, { sourceUri: sourceFile, assetRootUri: "assets" });
        const projectedHash = sha256(projected.programJson), geometryBound = { source: sourceBytes, assets: projected.assets, program: projected.programJson };
        const sourceNoop = await compilePpjWorkspace(geometryBound, { includePreviewScene: true });
        assert.deepEqual(sourceNoop.file, sourceBytes);
        const scenarios = [
          { name: "authored", workspace: geometryInput, receipt: author },
          { name: "source", workspace: geometryBound, receipt: sourceNoop },
        ];
        // Two independent source edits keep the custom geometry rather than
        // substituting a smaller rectangle to make the preview agree.
        if (geometry.name === "partial-rect") for (const edit of ["color", "frame"]) {
          const request = JSON.parse(Buffer.from(projected.programJson).toString("utf8"));
          if (edit === "color") for (const index of [0, 1]) request.pages[0].elements[0].style.fill.stops[index].color = "#FFFF00";
          else Object.assign(request.pages[0].elements[0].frame, { x: 240, width: 400 });
          const workspace = { ...geometryBound, program: Buffer.from(JSON.stringify(request)) };
          const requestFile = `radial-geometry-${geometry.name}-${edit}.ppj`;
          await writeFile(path.join(artifacts, requestFile), workspace.program, { flag: "wx" });
          const requestSha256 = sha256(workspace.program), receipt = await compilePpjWorkspace(workspace, { includePreviewScene: true });
          assert.equal(sha256(workspace.program), requestSha256);
          scenarios.push({ name: edit, workspace, receipt, request, requestFile, requestSha256 });
        }
        for (const scenario of scenarios) {
          const name = `geometry-${geometry.name}-${scenario.name}`, receipt = scenario.receipt;
          const candidateFile = `radial-${name}-candidate.pptx`;
          await writeFile(path.join(artifacts, candidateFile), receipt.file, { flag: "wx" });
          const painted = await savePaint(`radial-${name}`, receipt), native = createPpjSceneView(receipt).pages[0].nodes[0].native;
          assert.equal(native.customPaths.length, 1); assert.equal(native.gradientFill.kind, 2);
          const bounds = { x: scenario.name === "frame" ? 240 : 200, y: 140,
            width: scenario.name === "frame" ? 200 : geometry.width, height: geometry.height };
          const cx = bounds.x + bounds.width / 2, cy = bounds.y + bounds.height / 2;
          const svg = painted.pages[0].svg;
          const gradient = [...svg.matchAll(/<radialGradient\b([^>]*)>/g)]
            .map(g => Object.fromEntries([...g[1].matchAll(/([\w-]+)="([^"]*)"/g)].map(m => [m[1], m[2]])))
            .find(g => Number(g.cx) === cx && Number(g.cy) === cy);
          assert.ok(gradient, `${name}: actual path-bounds center`);
          assert.ok(Math.abs(Number(gradient.r) - Math.hypot(bounds.width, bounds.height) / 2) < 1e-9);
          const raster = await sharp(Buffer.from(svg)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
          const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * 4, (y * raster.info.width + x) * 4 + 4)];
          assert.deepEqual(pixel(cx, Math.floor(cy)), scenario.name === "color" ? [255, 255, 0, 255] : [255, 0, 0, 255]);
          assert.deepEqual(pixel(cx - bounds.width * .2, Math.floor(cy)), [0, 0, 0, 0], `${name}: transparent ring uses curve extrema`);
          assert.deepEqual(pixel(50, 360), [204, 85, 0, 255]);
          assert.match(svg, /HHHH/);
          assert.ok(!painted.diagnostics.some(d => ["preview.scene.paint.gradient-geometry", "preview.scene.paint.path"].includes(d.reason)));
          await assertProductionEntry(`radial-${name}`, scenario.workspace, receipt, painted);
          const record = { name, candidateFile, candidateSha256: sha256(receipt.file), sourceFile, sourceSha256: sourceDigest,
            actualPathBounds: bounds, nativeAndRgbaPixels: true, sourceNoop: scenario.name === "source" };
          if (scenario.request) {
            const fresh = await projectPptxToPpj(receipt.file, { sourceUri: candidateFile, assetRootUri: "assets" });
            // A new candidate gets new source-bound handles/revision hashes.
            // Compare the whole semantic page plus capability scopes and
            // ordered leaf kinds/values after checking candidate provenance;
            // native leaf IDs, like handles, are source-revision scoped.
            const comparablePage = (page, expectedSource) => {
              const copy = structuredClone(page);
              for (const owner of [copy, ...copy.elements]) {
                const ref = owner.nativeRef;
                assert.equal(ref.sourceSha256, expectedSource);
                assert.equal(ref.revision, `pptx-${expectedSource.slice(0, 16)}`);
                assert.match(ref.handle, /^nr-[0-9a-f]{64}$/);
                assert.match(ref.objectHash, /^[0-9a-f]{64}$/);
                assert.ok(ref.capabilities.every(c => c.expectedHash === ref.objectHash));
                for (const leaf of ref.leaves || []) assert.match(leaf.id, /^nl_[0-9a-f]{32}$/);
                owner.nativeRef = { capabilities: ref.capabilities.map(({ expectedHash, ...scope }) => scope),
                  leaves: ref.leaves?.map(({ expectedHash, id, ...leaf }) => leaf) };
              }
              return copy;
            };
            const expectedPage = comparablePage(scenario.request.pages[0], sourceDigest);
            if (scenario.name === "frame") for (const leaf of expectedPage.elements[0].nativeRef.leaves) {
              if (leaf.kind === "leftEmu") leaf.value = 240 * 12700;
              if (leaf.kind === "widthEmu") leaf.value = 400 * 12700;
            }
            assert.deepEqual(comparablePage(JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0], sha256(receipt.file)), expectedPage);
            const freshNoop = await compilePpjWorkspace({ source: receipt.file, assets: fresh.assets, program: fresh.programJson });
            assert.deepEqual(freshNoop.file, receipt.file);
            const originalZip = await JSZip.loadAsync(sourceBytes), candidateZip = await JSZip.loadAsync(receipt.file), changedParts = [];
            assert.deepEqual(Object.keys(candidateZip.files).sort(), Object.keys(originalZip.files).sort());
            for (const part of Object.keys(originalZip.files)) if (!originalZip.files[part].dir &&
              !Buffer.from(await originalZip.file(part).async("uint8array")).equals(Buffer.from(await candidateZip.file(part).async("uint8array")))) changedParts.push(part);
            assert.deepEqual(changedParts, ["ppt/slides/slide1.xml"]);
            Object.assign(record, { reprojection: true, changedParts, requestFile: scenario.requestFile, requestSha256: scenario.requestSha256 });
          }
          radialGeometryCases.push(record);
        }
        assert.equal(sha256(geometryInput.program), geometryInputHash);
        assert.equal(sha256(projected.programJson), projectedHash); assert.equal(sha256(sourceBytes), sourceDigest);
      } catch (error) {
        radialGradientFailures.push({ name: `geometry-${geometry.name}`, code: error.code, message: error.message });
      }
    }
    assert.equal(sha256(input.program), inputHash); assert.equal(sha256(projection.programJson), projectionHash); assert.equal(sha256(source), sourceHash);
  } catch (error) {
    radialGradientFailures.push({ name: "authored/source foundation", code: error.code, message: error.message });
    console.error(`Radial gradient foundation failed: ${error.message}`);
  }
  // One gradient consumer per fixture makes background compositing explicit:
  // shapes/cells sit on white; a page gradient keeps its own transparent alpha.
  for (const owner of ["shape", "table", "background"]) for (const kind of ["linear", "radial"]) {
    const fixtureName = `${owner}-${kind}`;
    try {
      const fill = { type: "gradient", kind, ...(kind === "linear" ? { angle: 0 } : {}), stops: [
        { offset: 0, color: "#C02010", opacity: .25 }, { offset: 1, color: "#1050E0", opacity: .75 },
      ] };
      const program = structuredClone(pairBase), page = program.pages[0], frame = { x: 128, y: 96, width: 256, height: 128 };
      page.background = owner === "background" ? fill : { type: "solid", color: "#FFFFFF" };
      page.elements = owner === "background" ? [] : owner === "shape" ? [{ id: "brightness-shape", type: "shape", frame,
        geometry: { kind: "preset", preset: "rect" }, style: { fill } }] : [{ id: "brightness-table", type: "table", frame,
        columns: [{ id: "column", width: frame.width }], rows: [{ id: "row", height: frame.height, cells: [{ id: "cell", text: "", fill }] }] }];
      page.elements.push({ id: "brightness-control", type: "shape", frame: { x: 32, y: 40, width: 20, height: 20 },
        geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } });
      const input = { ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)) }, inputHash = sha256(input.program);
      const authored = await compilePpjWorkspace(input, { includePreviewScene: true });
      const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source), sourceFile = `brightness-${fixtureName}-source.pptx`;
      await writeFile(path.join(artifacts, sourceFile), source, { flag: "wx" });
      const projection = await projectPptxToPpj(source, { sourceUri: sourceFile, assetRootUri: "assets" });
      const projectionHash = sha256(projection.programJson), bound = { source, assets: projection.assets, program: projection.programJson };
      const sourceNoop = await compilePpjWorkspace(bound, { includePreviewScene: true });
      assert.deepEqual(sourceNoop.file, source);
      const fillAt = p => owner === "background" ? p.pages[0].background : owner === "shape"
        ? p.pages[0].elements[0].style.fill : p.pages[0].elements[0].rows[0].cells[0].fill;
      async function brightnessPixels(name, compiled, requestedFill) {
        const painted = await savePaint(`brightness-${name}`, compiled), view = createPpjSceneView(compiled);
        const native = owner === "background" ? view.pages[0].native.background.gradientFill : owner === "shape"
          ? view.pages[0].nodes[0].native.gradientFill : view.pages[0].nodes[0].native.rows[0].cells[0].fill.kind.value;
        assert.equal(native.kind, kind === "linear" ? 1 : 2);
        assert.equal(native.stops.length, requestedFill.stops.length, "SVG samples must not be written into the PPTX");
        for (let i = 0; i < native.stops.length; i++) {
          const expected = requestedFill.stops[i], actual = native.stops[i];
          assert.equal(actual.positionThousandthPercent, Math.round(expected.offset * 100000));
          assert.equal(actual.colorRgb, expected.color.slice(1).toUpperCase());
          assert.equal(actual.opacityThousandthPercent, Math.round(expected.opacity * 100000));
        }
        const sampled = requestedFill.stops.at(-1).offset === 1;
        const svg = painted.pages[0].svg;
        assert.equal((svg.match(/<stop /g) || []).length, sampled ? 32 * (native.stops.length - 1) + 1 : native.stops.length);
        const diagnostic = painted.diagnostics.filter(d => d.reason === "preview.scene.paint.gradient-interpolation");
        assert.equal(diagnostic.length, sampled ? 1 : 0);
        assert.ok(diagnostic.every(d => d.status === "partial" && d.valueSummary.includes("32 SVG segments")));
        const { data, info } = await sharp(Buffer.from(svg)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
        const pixel = (x, y) => [...data.subarray((y * info.width + x) * 4, (y * info.width + x) * 4 + 4)];
        const f = owner === "background" ? { x: 0, y: 0, width: 960, height: 540 } : frame;
        const cx = f.x + f.width / 2, cy = f.y + f.height / 2, radius = Math.hypot(f.width, f.height) / 2;
        const colors = requestedFill.stops.map(s => s.color.slice(1).match(/../g).map(h => parseInt(h, 16)));
        for (const fraction of [0, .125, .25, .5, .75]) {
          const x = kind === "linear" ? Math.floor(f.x + f.width * Math.max(fraction, .015625)) : Math.floor(cx + radius * fraction);
          const y = Math.floor(cy), position = kind === "linear" ? (x + .5 - f.x) / f.width : Math.hypot(x + .5 - cx, y + .5 - cy) / radius;
          let interval = requestedFill.stops.findIndex(s => s.offset >= position) - 1;
          if (interval < 0) interval = requestedFill.stops.length - 2;
          const a = requestedFill.stops[interval], b = requestedFill.stops[interval + 1];
          const t = Math.max(0, Math.min(1, (position - a.offset) / (b.offset - a.offset)));
          const alpha = a.opacity + (b.opacity - a.opacity) * t;
          const expected = colors[interval].map((v, c) => sampled
            ? Math.max(v, colors[interval + 1][c]) - Math.abs(v - colors[interval + 1][c]) * (v < colors[interval + 1][c] ? 1 - t : t) ** (15 / 8)
            : v + (colors[interval + 1][c] - v) * t);
          const actual = pixel(x, y), destinationAlpha = owner === "background" ? alpha : 1;
          assert.ok(Math.abs(actual[3] - destinationAlpha * 255) <= 1, `${name}: linear alpha at ${x},${y}`);
          for (let c = 0; c < 3; c++) {
            const premultiplied = expected[c] * alpha + (owner === "background" ? 0 : 255 * (1 - alpha));
            assert.ok(Math.abs(actual[c] * actual[3] / 255 - premultiplied) <= 2,
              `${name}: composited RGB channel ${c} at ${x},${y}: ${actual} != ${expected}, alpha=${alpha}`);
          }
          if (alpha === 0) assert.deepEqual(actual, owner === "background" ? [0, 0, 0, 0] : [255, 255, 255, 255]);
        }
        assert.deepEqual(pixel(40, 48), [204, 85, 0, 255]);
        return painted;
      }
      for (const [name, workspace, compiled] of [["authored", input, authored], ["source", bound, sourceNoop]]) {
        const fullName = `${fixtureName}-${name}`, painted = await brightnessPixels(fullName, compiled, fill);
        await assertProductionEntry(`brightness-${fullName}`, workspace, compiled, painted);
        brightnessCases.push({ name: fullName, nativeAndCompositedRgbaPixels: true, sourceNoop: name === "source", sourceFile, sourceSha256: sourceHash,
          candidateSha256: sha256(compiled.file), originalStopCount: 2, svgStopCount: 33 });
      }
      for (const edit of ["reverse", "zero-alpha", "three-color", "non-endpoint"]) {
        const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8")), editedFill = fillAt(request);
        if (edit === "reverse") [editedFill.stops[0].color, editedFill.stops[1].color] = [editedFill.stops[1].color, editedFill.stops[0].color];
        if (edit === "zero-alpha") for (const stop of editedFill.stops) stop.opacity = 0;
        if (edit === "three-color") editedFill.stops = [{ ...editedFill.stops[0] }, { ...editedFill.stops[1], offset: .25, opacity: 1 },
          { ...editedFill.stops[0], offset: 1, opacity: .5 }];
        if (edit === "non-endpoint") editedFill.stops[1].offset = .9;
        const name = `${fixtureName}-${edit}`, requestFile = `brightness-${name}.ppj`;
        const edited = { ...bound, program: Buffer.from(JSON.stringify(request)) }, requestHash = sha256(edited.program);
        await writeFile(path.join(artifacts, requestFile), edited.program, { flag: "wx" });
        const evidence = { name, stage: "compile", sourceFile, sourceSha256: sourceHash, requestFile, requestSha256: requestHash };
        try {
          const candidate = await compilePpjWorkspace(edited, { includePreviewScene: true }), candidateFile = `brightness-${name}-candidate.pptx`;
          await writeFile(path.join(artifacts, candidateFile), candidate.file, { flag: "wx" });
          Object.assign(evidence, { stage: "paint", candidateFile, candidateSha256: sha256(candidate.file) });
          const painted = await brightnessPixels(name, candidate, editedFill);
          await assertProductionEntry(`brightness-${name}`, edited, candidate, painted);
          evidence.stage = "reprojection";
          const fresh = await projectPptxToPpj(candidate.file, { sourceUri: candidateFile, assetRootUri: "assets" });
          const reprojected = JSON.parse(Buffer.from(fresh.programJson).toString("utf8"));
          const reprojectedFile = `brightness-${name}-reprojected.ppj`;
          await writeFile(path.join(artifacts, reprojectedFile), fresh.programJson, { flag: "wx" });
          if (owner === "shape" && edit === "zero-alpha") {
            // The established projector factors equal paint opacity on a
            // textless shape into compositing.opacity. Check that exact
            // normalization, not an arbitrary omitted-alpha exception.
            assert.deepEqual(reprojected.pages[0].elements[0].compositing, { opacity: 0 });
            const normalizedFill = { ...editedFill, stops: editedFill.stops.map(({ opacity, ...stop }) => stop) };
            assert.deepEqual(fillAt(reprojected), normalizedFill);
            const freshNoop = await compilePpjWorkspace({ source: candidate.file, assets: fresh.assets, program: fresh.programJson }, { includePreviewScene: true });
            assert.deepEqual(freshNoop.file, candidate.file);
            await brightnessPixels(`${name}-reprojected`, freshNoop, editedFill);
            evidence.compoundZeroOpacityPreserved = true;
          } else assert.deepEqual(fillAt(reprojected), editedFill);
          const originalZip = await JSZip.loadAsync(source), candidateZip = await JSZip.loadAsync(candidate.file), changedParts = [];
          assert.deepEqual(Object.keys(candidateZip.files).sort(), Object.keys(originalZip.files).sort());
          for (const part of Object.keys(originalZip.files)) if (!originalZip.files[part].dir &&
            !Buffer.from(await originalZip.file(part).async("uint8array")).equals(Buffer.from(await candidateZip.file(part).async("uint8array")))) changedParts.push(part);
          assert.deepEqual(changedParts, ["ppt/slides/slide1.xml"]);
          brightnessCases.push({ ...evidence, stage: "verified", nativeAndCompositedRgbaPixels: true, reprojection: true,
            reprojectedFile, reprojectedSha256: sha256(fresh.programJson), changedParts,
            originalStopCount: editedFill.stops.length, svgStopCount: edit === "non-endpoint" ? 2 : edit === "three-color" ? 65 : 33 });
        } catch (error) { brightnessFailures.push({ ...evidence, code: error.code, message: error.message }); }
        assert.equal(sha256(source), sourceHash); assert.equal(sha256(edited.program), requestHash);
      }
      assert.equal(sha256(input.program), inputHash); assert.equal(sha256(projection.programJson), projectionHash);
    } catch (error) { brightnessFailures.push({ name: fixtureName, code: error.code, message: error.message }); }
  }
  const mediaPosterCases = [];
  // Container header only: enough for the codec's embedded-media contract,
  // deliberately not a playable clip or playback acceptance fixture.
  const mediaBytes = Buffer.from("000000186674797069736F6D0000020069736F6D6D703431", "hex");
  for (const color of ["CC5500", "0066CC"]) {
    const poster = Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="100"><rect width="100" height="100" fill="#${color}"/></svg>`);
    const program = structuredClone(pairBase);
    program.pages[0].elements = [{ id: "media", type: "media", mediaType: "video", asset: "clip", posterAsset: "poster",
      frame: { x: 100, y: 100, width: 200, height: 100 }, playback: { trigger: "onSlideStart" } }];
    program.pages[0].readingOrder = ["media"];
    const assets = [{ id: "clip", uri: "assets/clip.mp4", mimeType: "video/mp4", data: mediaBytes },
      { id: "poster", uri: "assets/poster.svg", mimeType: "image/svg+xml", data: poster }].map(asset => ({ ...asset, sha256: sha256(asset.data) }));
    program.assets = assets.map(({ data, ...asset }) => ({ ...asset, rights: { status: "internal" },
      accessibility: { decorative: false, description: `Synthetic ${asset.id} for static-poster checks` } }));
    const input = { ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)), assets };
    const inputHash = sha256(input.program), authored = await compilePpjWorkspace(input, { includePreviewScene: true });
    const native = authored.previewScene.presentation.slides[0].elements[0];
    assert.equal(native.content.case, "media");
    assert.equal(native.content.value.mediaType, "video");
    const painted = await savePaint(`media-poster-${color}`, authored);
    assert.match(painted.pages[0].svg, /data-officekit-media="static-poster"/);
    assert.match(painted.pages[0].svg, /STATIC VIDEO POSTER/);
    assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.media-static"));
    assert.ok(painted.diagnostics.some(d => d.scenePath.endsWith("media.playbackTrigger")));
    const raster = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * raster.info.channels, (y * raster.info.width + x) * raster.info.channels + 3)];
    assert.deepEqual(pixel(150, 140), color.match(/../g).map(value => parseInt(value, 16)));
    assert.deepEqual(pixel(250, 140), [255, 255, 255], "Transparent poster half keeps the page background visible");
    assert.deepEqual(pixel(290, 190), [146, 64, 14], "Static-poster warning is visible in the exported pixels");
    const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    const projected = await projectPptxToPpj(source, { sourceUri: "media-poster.pptx", assetRootUri: "assets" });
    const noop = await compilePpjWorkspace({ program: projected.programJson, source, assets: projected.assets }, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    const sourcePaint = await savePaint(`media-poster-${color}-source`, noop);
    assert.equal(noop.previewScene.presentation.slides[0].elements[0].content.case, "opaque");
    assert.ok(!sourcePaint.pages[0].svg.includes('data-officekit-media="static-poster"'));
    assert.ok(sourcePaint.diagnostics.some(d => d.reason === "preview.scene.paint.content"));
    assert.equal(sha256(source), sourceHash);
    assert.equal(sha256(input.program), inputHash);
    assert.equal(sha256(assets[0].data), assets[0].sha256);
    assert.equal(sha256(assets[1].data), assets[1].sha256);
    mediaPosterCases.push({ color, authoredPixels: true, transparency: true, visibleStaticWarning: true,
      sourceNoopBytes: true, sourceOpaqueRetained: true, sourceSha256: sourceHash, candidateSha256: sha256(authored.file) });
  }
  const capitalizationCases = [];
  const capsPixels = async (name, receipt) => {
    const painted = await savePaint(name, receipt);
    return sharp(Buffer.from(painted.pages[0].svg)).extract({ left: 100, top: 100, width: 300, height: 100 }).removeAlpha().raw().toBuffer();
  };
  for (const caps of ["none", "all"]) {
    const program = structuredClone(formatProgram), run = program.pages[0].elements[0].text.paragraphs[0].runs[0];
    run.text = "Hello abc";
    run.style.capitalization = caps;
    const programBefore = JSON.stringify(program), reference = structuredClone(program);
    reference.pages[0].elements[0].text.paragraphs[0].runs[0].text = caps === "all" ? "HELLO ABC" : "Hello abc";
    reference.pages[0].elements[0].text.paragraphs[0].runs[0].style.capitalization = "none";
    const expected = await capsPixels(`caps-${caps}-reference`, await compileFormat(reference));
    const authored = await compileFormat(program), source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    const projected = await projectPptxToPpj(source, { sourceUri: "caps.pptx", assetRootUri: "assets" });
    const input = { program: projected.programJson, source, assets: projected.assets };
    const noop = await compilePpjWorkspace(input, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    for (const [origin, receipt] of [["authored", authored], ["source", noop]]) {
      const nativeRun = receipt.previewScene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0].runs[0];
      assert.equal(nativeRun.fontCaps, caps);
      assert.equal(nativeRun.content.value, "Hello abc");
      assert.deepEqual(await capsPixels(`caps-${caps}-${origin}`, receipt), expected);
      capitalizationCases.push({ caps, origin, literalPreserved: true, pixels: true, candidateSha256: sha256(receipt.file) });
    }
    if (caps === "all") {
      const request = JSON.parse(Buffer.from(projected.programJson).toString("utf8"));
      const leaf = request.pages[0].elements[0].nativeRef.leaves.find(item => item.kind === "fontCaps");
      assert.ok(leaf);
      leaf.value = "none";
      const candidate = await compilePpjWorkspace({ ...input, program: Buffer.from(JSON.stringify(request)) }, { includePreviewScene: true });
      reference.pages[0].elements[0].text.paragraphs[0].runs[0].text = "Hello abc";
      assert.deepEqual(await capsPixels("caps-source-edited", candidate), await capsPixels("caps-source-edited-reference", await compileFormat(reference)));
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: "caps-edited.pptx", assetRootUri: "assets" });
      const freshRun = JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0].elements[0].text.paragraphs[0].runs[0];
      assert.equal(freshRun.text, "Hello abc");
      assert.equal(freshRun.style.capitalization, "none");
      const before = await JSZip.loadAsync(source), after = await JSZip.loadAsync(candidate.file), changed = [];
      assert.deepEqual(Object.keys(after.files).sort(), Object.keys(before.files).sort());
      for (const name of Object.keys(before.files)) if (!before.files[name].dir &&
        !Buffer.from(await before.file(name).async("uint8array")).equals(Buffer.from(await after.file(name).async("uint8array")))) changed.push(name);
      assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
      capitalizationCases.push({ caps: "none", origin: "source-edit", literalPreserved: true, pixels: true, changedParts: changed });
    }
    assert.equal(sha256(source), sourceHash);
    assert.equal(JSON.stringify(program), programBefore);
  }
  const characterBulletCases = [];
  const bulletProgram = structuredClone(formatProgram);
  bulletProgram.pages[0].elements[0].text.paragraphs = [{ style: { indent: 30, hanging: 20,
    bullet: { type: "character", character: "●", fontFamily: "DejaVu Sans", color: "#CC5500", size: 20 } },
    runs: [{ text: "HH\nHH", style: { size: 20, color: "#000000" } }] }];
  const bulletProgramHash = sha256(Buffer.from(JSON.stringify(bulletProgram)));
  const bulletAuthored = await compileFormat(bulletProgram), bulletSource = await withoutAuthoredSnapshot(bulletAuthored.file), bulletSourceHash = sha256(bulletSource);
  const bulletProjection = await projectPptxToPpj(bulletSource, { sourceUri: "character-bullet.pptx", assetRootUri: "assets" });
  const bulletWorkspace = { program: bulletProjection.programJson, source: bulletSource, assets: bulletProjection.assets };
  const bulletNoop = await compilePpjWorkspace(bulletWorkspace, { includePreviewScene: true });
  assert.deepEqual(bulletNoop.file, bulletSource);
  async function bulletPixels(origin, receipt, character = "●", color = "CC5500") {
    const native = receipt.previewScene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0];
    assert.equal(native.bullet.value, character);
    assert.equal(native.bulletColor.value, color);
    assert.equal(native.bulletSize.value, 20);
    const painted = await savePaint(`bullet-${origin}`, receipt), svg = painted.pages[0].svg;
    assert.equal(svg.match(/data-officekit-bullet="character"/g)?.length, 1);
    assert.match(svg, /data-officekit-bullet="character" x="117.19999999999999" y="123.6"/);
    assert.ok(svg.includes('<text x="137.2" y="123.6"'));
    assert.ok(svg.includes('<text x="137.2" y="147.6"'));
    const raster = await sharp(Buffer.from(svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const rgb = color.match(/../g).map(value => parseInt(value, 16));
    let first = 0, next = 0;
    for (let y = 100; y < 160; y++) for (let x = 110; x < 137; x++) {
      const offset = (y * raster.info.width + x) * raster.info.channels;
      if (rgb.every((value, index) => raster.data[offset + index] === value)) {
        if (y < 130) first++; else next++;
      }
    }
    assert.ok(first > 0, "Bullet must have actual colored glyph pixels");
    assert.equal(next, 0, "Explicit continuation line must not repeat the bullet");
    characterBulletCases.push({ origin, character, color, glyphPixels: first, candidateSha256: sha256(receipt.file) });
    return sharp(Buffer.from(svg)).extract({ left: 100, top: 100, width: 300, height: 100 }).removeAlpha().raw().toBuffer();
  }
  assert.deepEqual(await bulletPixels("authored", bulletAuthored), await bulletPixels("source", bulletNoop));
  const bulletEdit = JSON.parse(Buffer.from(bulletProjection.programJson).toString("utf8"));
  for (const [kind, value] of [["paragraphBulletCharacter", "■"], ["paragraphBulletColorRgb", "0066CC"]]) {
    const leaf = bulletEdit.pages[0].elements[0].nativeRef.leaves.find(item => item.kind === kind);
    assert.ok(leaf, `Source must issue ${kind}`);
    leaf.value = value;
  }
  const bulletCandidate = await compilePpjWorkspace({ ...bulletWorkspace, program: Buffer.from(JSON.stringify(bulletEdit)) }, { includePreviewScene: true });
  const bulletExpected = structuredClone(bulletProgram);
  Object.assign(bulletExpected.pages[0].elements[0].text.paragraphs[0].style.bullet, { character: "■", color: "#0066CC" });
  assert.deepEqual(await bulletPixels("source-edit", bulletCandidate, "■", "0066CC"),
    await bulletPixels("explicit-reference", await compileFormat(bulletExpected), "■", "0066CC"));
  const bulletFresh = await projectPptxToPpj(bulletCandidate.file, { sourceUri: "character-bullet-edited.pptx", assetRootUri: "assets" });
  const freshBullet = JSON.parse(Buffer.from(bulletFresh.programJson).toString("utf8")).pages[0].elements[0].text.paragraphs[0].style.bullet;
  assert.equal(freshBullet.character, "■");
  const bulletBeforeZip = await JSZip.loadAsync(bulletSource), bulletAfterZip = await JSZip.loadAsync(bulletCandidate.file), bulletChangedParts = [];
  assert.deepEqual(Object.keys(bulletAfterZip.files).sort(), Object.keys(bulletBeforeZip.files).sort());
  for (const name of Object.keys(bulletBeforeZip.files)) if (!bulletBeforeZip.files[name].dir &&
    !Buffer.from(await bulletBeforeZip.file(name).async("uint8array")).equals(Buffer.from(await bulletAfterZip.file(name).async("uint8array")))) bulletChangedParts.push(name);
  assert.deepEqual(bulletChangedParts, ["ppt/slides/slide1.xml"]);
  assert.equal(sha256(bulletSource), bulletSourceHash);
  assert.equal(sha256(Buffer.from(JSON.stringify(bulletProgram))), bulletProgramHash);
  const smallTextProgram = structuredClone(formatProgram);
  smallTextProgram.pages[0].elements[0].text.paragraphs[0].runs[0].style.size = 8;
  const smallTextAuthored = await compileFormat(smallTextProgram);
  const smallTextSource = await withoutAuthoredSnapshot(smallTextAuthored.file), smallTextHash = sha256(smallTextSource);
  const smallTextProjection = await projectPptxToPpj(smallTextSource, { sourceUri: "small-text.pptx", assetRootUri: "assets" });
  const smallTextNoop = await compilePpjWorkspace({ program: smallTextProjection.programJson, source: smallTextSource, assets: smallTextProjection.assets }, { includePreviewScene: true });
  assert.deepEqual(smallTextNoop.file, smallTextSource);
  for (const [name, receipt] of [["authored", smallTextAuthored], ["source", smallTextNoop]]) {
    const painted = await savePaint(`small-text-${name}`, receipt);
    assert.match(painted.pages[0].svg, /<text x="107\.2" y="111\.6"/);
    const raster = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const darkCount = (start, end) => {
      let count = 0;
      for (let y = start; y < end; y++) for (let x = 100; x < 200; x++) {
        const offset = (y * raster.info.width + x) * raster.info.channels;
        if ([...raster.data.subarray(offset, offset + 3)].every(v => v < 200)) count++;
      }
      return count;
    };
    assert.ok(darkCount(103, 113) > 0, "8pt glyphs must use the 8pt baseline");
    assert.equal(darkCount(120, 136), 0, "default 18pt must not displace an explicitly small run");
  }
  assert.equal(sha256(smallTextSource), smallTextHash);
  for (const alignment of ["right", "center"]) {
    const inkBounds = [];
    for (const right of [0, 30]) {
      const program = structuredClone(formatProgram);
      program.pages[0].elements[0].style = { margins: { left: 0, right, top: 0, bottom: 0 } };
      program.pages[0].elements[0].text.paragraphs[0].style = { alignment };
      const receipt = await compileFormat(program), painted = await savePaint(`inset-${alignment}-${right}`, receipt);
      const expectedX = alignment === "right" ? 400 - right : (500 - right) / 2;
      assert.ok(painted.pages[0].svg.includes(`<text x="${expectedX}"`));
      assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && /\.(leftInsetEmu|rightInsetEmu|topInsetEmu)$/.test(d.scenePath)));
      const raster = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
      let min = Infinity, max = -Infinity, pixels = 0;
      for (let y = 100; y < 190; y++) for (let x = 100; x < 410; x++) {
        const offset = (y * raster.info.width + x) * raster.info.channels;
        if ([...raster.data.subarray(offset, offset + 3)].every(v => v < 80)) { min = Math.min(min, x); max = Math.max(max, x); pixels++; }
      }
      assert.ok(pixels > 0);
      inkBounds.push({ min, max, pixels });
    }
    const shift = alignment === "right" ? -30 : -15;
    assert.equal(inkBounds[1].min - inkBounds[0].min, shift);
    assert.equal(inkBounds[1].max - inkBounds[0].max, shift);
    assert.equal(inkBounds[1].pixels, inkBounds[0].pixels);
  }
  const textAnchorCases = [], textAnchorEdits = [], textAnchorFailures = [];
  for (const bottom of [10, 30]) for (const verticalAlignment of ["top", "middle", "bottom"]) {
    const program = structuredClone(formatProgram);
    program.pages[0].elements[0].style = { verticalAlignment, margins: { left: 0, right: 0, top: 10, bottom } };
    const originalProgram = JSON.stringify(program), authored = await compileFormat(program);
    const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    const projected = await projectPptxToPpj(source, { sourceUri: "text-anchor.pptx", assetRootUri: "assets" });
    const noop = await compilePpjWorkspace({ program: projected.programJson, source, assets: projected.assets }, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    const anchor = verticalAlignment === "middle" ? "center" : verticalAlignment;
    const shift = anchor === "top" ? 0 : (100 - 10 - bottom - 48) / (anchor === "center" ? 2 : 1);
    for (const [origin, receipt] of [["authored", authored], ["source", noop]]) {
      const shape = receipt.previewScene.presentation.slides[0].elements[0].content.value;
      assert.equal(shape.textBody.bodyProperties.anchor.value, anchor);
      assert.equal(shape.textBody.bodyProperties.bottomInset.value, BigInt(bottom * 12700));
      const painted = await savePaint(`text-anchor-${origin}-${anchor}-${bottom}`, receipt);
      assert.ok(painted.pages[0].svg.includes(`data-officekit-text-anchor="${anchor}" transform="translate(0 ${shift})"`));
      assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-layout"));
      assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && /\.(verticalAnchor|bottomInsetEmu)$/.test(d.scenePath)));
      const raster = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
      let min = Infinity, max = -Infinity, pixels = 0;
      for (let y = 100; y < 200; y++) for (let x = 100; x < 390; x++) {
        const offset = (y * raster.info.width + x) * raster.info.channels;
        if ([...raster.data.subarray(offset, offset + 3)].every(v => v < 80)) { min = Math.min(min, y); max = Math.max(max, y); pixels++; }
      }
      assert.ok(pixels > 0);
      textAnchorCases.push({ origin, anchor, bottom, shift, min, max, pixels, sourceSha256: sourceHash, candidateSha256: sha256(receipt.file) });
    }
    if (bottom === 30 && anchor === "bottom") {
      await writeFile(path.join(artifacts, "text-anchor-edit-source.pptx"), source, { flag: "wx" });
      const originalZip = await JSZip.loadAsync(source);
      for (const operation of ["middle", "top", "zero-bottom", "delete-anchor", "delete-bottom", "delete-both"]) {
        try {
          // Every edit starts from the original package, never the previous candidate.
          const projection = await projectPptxToPpj(source, { sourceUri: "text-anchor.pptx", assetRootUri: "assets" });
          const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8"));
          const styleOf = p => {
            const element = p.pages[0].elements[0];
            return element.type === "text" ? element.style : element.textStyle;
          };
          const requestedStyle = styleOf(request), expectedProgram = structuredClone(program);
          assert.equal(requestedStyle.verticalAlignment, "bottom");
          assert.equal(requestedStyle.margins.bottom, 30);
          for (const style of [requestedStyle, expectedProgram.pages[0].elements[0].style]) {
            if (operation === "middle" || operation === "top") style.verticalAlignment = operation;
            if (operation === "zero-bottom") style.margins.bottom = 0;
            if (operation === "delete-anchor" || operation === "delete-both") delete style.verticalAlignment;
            if (operation === "delete-bottom" || operation === "delete-both") delete style.margins.bottom;
          }
          const requestBytes = Buffer.from(JSON.stringify(request)), requestHash = sha256(requestBytes);
          const candidate = await compilePpjWorkspace({ program: requestBytes, source, assets: projection.assets }, { includePreviewScene: true });
          await writeFile(path.join(artifacts, `text-anchor-${operation}.pptx`), candidate.file, { flag: "wx" });
          const fresh = await projectPptxToPpj(candidate.file, { sourceUri: "text-anchor-candidate.pptx", assetRootUri: "assets" });
          await writeFile(path.join(artifacts, `text-anchor-${operation}.reprojected.ppj`), fresh.programJson, { flag: "wx" });
          const actualStyle = styleOf(JSON.parse(Buffer.from(fresh.programJson).toString("utf8")));
          assert.equal(Object.hasOwn(actualStyle, "verticalAlignment"), Object.hasOwn(requestedStyle, "verticalAlignment"), `${operation}: anchor presence must match request`);
          assert.equal(actualStyle.verticalAlignment, requestedStyle.verticalAlignment);
          assert.equal(Object.hasOwn(actualStyle.margins, "bottom"), Object.hasOwn(requestedStyle.margins, "bottom"), `${operation}: bottom inset presence must match request`);
          assert.equal(actualStyle.margins.bottom, requestedStyle.margins.bottom);
          const properties = candidate.previewScene.presentation.slides[0].elements[0].content.value.textBody.bodyProperties;
          assert.equal(properties.anchor.case, requestedStyle.verticalAlignment === undefined ? undefined : "verticalAnchor");
          assert.equal(properties.anchor.value, requestedStyle.verticalAlignment === "middle" ? "center" : requestedStyle.verticalAlignment);
          assert.equal(properties.bottomInset.case, requestedStyle.margins.bottom === undefined ? undefined : "bottomInsetEmu");
          assert.equal(properties.bottomInset.value, requestedStyle.margins.bottom === undefined ? undefined : BigInt(requestedStyle.margins.bottom * 12700));
          const candidateZip = await JSZip.loadAsync(candidate.file);
          assert.deepEqual(Object.keys(candidateZip.files).sort(), Object.keys(originalZip.files).sort());
          const changedParts = [];
          for (const name of Object.keys(originalZip.files)) if (!originalZip.files[name].dir &&
            !Buffer.from(await originalZip.file(name).async("uint8array")).equals(Buffer.from(await candidateZip.file(name).async("uint8array")))) changedParts.push(name);
          assert.deepEqual(changedParts, ["ppt/slides/slide1.xml"]);
          const xml = await candidateZip.file("ppt/slides/slide1.xml").async("string");
          const bodyPr = xml.match(/<a:bodyPr\b[^>]*>/)?.[0];
          assert.ok(bodyPr);
          assert.equal(/\banchor=/.test(bodyPr), requestedStyle.verticalAlignment !== undefined);
          assert.equal(/\bbIns=/.test(bodyPr), requestedStyle.margins.bottom !== undefined);
          if (operation === "top") assert.match(bodyPr, /\banchor="t"/);
          if (operation === "zero-bottom") assert.match(bodyPr, /\bbIns="0"/);
          const painted = await savePaint(`text-anchor-${operation}`, candidate);
          const explicit = await savePaint(`text-anchor-${operation}-explicit`, await compileFormat(expectedProgram));
          // Compare the actual glyph region, not the source/author warning banners.
          const pixels = async p => sharp(Buffer.from(p.pages[0].svg)).extract({ left: 100, top: 100, width: 300, height: 100 })
            .removeAlpha().raw().toBuffer();
          assert.deepEqual(await pixels(painted), await pixels(explicit));
          assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-layout"));
          assert.equal(sha256(requestBytes), requestHash);
          assert.equal(sha256(source), sourceHash);
          textAnchorEdits.push({ operation, sourceSha256: sourceHash, candidateSha256: sha256(candidate.file),
            actualStyle, changedParts, actualGlyphRegionEqual: true });
        } catch (error) {
          textAnchorFailures.push({ operation, message: error.message, code: error.code ?? null });
          console.error(`Text anchor ${operation} failed: ${error.message}`);
        }
      }
    }
    assert.equal(sha256(source), sourceHash);
    assert.equal(JSON.stringify(program), originalProgram);
  }
  for (const row of textAnchorCases) {
    const top = textAnchorCases.find(c => c.origin === row.origin && c.anchor === "top" && c.bottom === row.bottom);
    assert.equal(row.min - top.min, row.shift);
    assert.equal(row.max - top.max, row.shift);
    assert.equal(row.pixels, top.pixels);
  }
  const textRotationCases = [], textRotationFailures = [], textRotationOpaqueCases = [];
  const rotationFrame = { x: 128, y: 160, width: 256, height: 128 };
  const rotationProgram = (owner, angle, outer = 0, flip = false) => {
    const program = structuredClone(pairBase), frame = { ...rotationFrame, rotation: outer, flipH: flip };
    const style = { rotation: angle, verticalAlignment: "middle", margins: { left: 8, right: 12, top: 10, bottom: 14 } };
    const text = { paragraphs: [{ runs: [{ text: "F0", style: { size: 24, color: "#0055CC" } }, { break: true },
      { text: "IL", style: { size: 24, color: "#0055CC" } }] }] };
    let element;
    if (owner === "text") element = { id: "rotation", type: "text", frame, style, text };
    if (owner === "shape") element = { id: "rotation", type: "shape", frame,
      geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#F5F0E0" } }, textStyle: style, text };
    if (owner === "table") element = { id: "rotation", type: "table", frame: { ...rotationFrame },
      columns: [{ id: "c", width: 256 }], rows: [{ id: "r", height: 128, cells: [{ id: "cell",
        fill: { type: "solid", color: "#F5F0E0" }, text: { ...text, style } }] }] };
    program.pages[0].elements = [element, { id: "control", type: "shape", frame: { x: 40, y: 80, width: 24, height: 16 },
      geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC6600" } } }];
    return program;
  };
  const rotationStyle = (program, owner) => {
    const element = program.pages[0].elements[0];
    return owner === "table" ? element.rows[0].cells[0].text.style : element.type === "text" ? element.style : element.textStyle;
  };
  const rotationBody = (receipt, owner) => {
    const element = receipt.previewScene.presentation.slides[0].elements[0];
    assert.equal(element.content.case, owner === "table" ? "table" : "shape", "source owner must retain its modeled kind before accessing text");
    const native = element.content.value;
    return owner === "table" ? native.rows[0].cells[0].textBody : native.textBody;
  };
  async function rotationInk(painted) {
    const { data, info } = await sharp(Buffer.from(painted.pages[0].svg)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
    const points = [];
    for (let y = 24; y < info.height; y++) for (let x = 0; x < info.width; x++) {
      const i = (y * info.width + x) * info.channels;
      if (data[i] < 60 && data[i + 1] >= 50 && data[i + 1] < 150 && data[i + 2] > 140 && data[i + 3] > 200) points.push([x, y]);
    }
    assert.ok(points.length > 100, "asymmetric text must actually have visible glyph pixels");
    const i = (84 * info.width + 44) * info.channels;
    assert.deepEqual([...data.subarray(i, i + 4)], [204, 102, 0, 255], "unrelated foreground must not rotate");
    return { points, data, info };
  }
  const rotationBounds = points => ({ left: Math.min(...points.map(p => p[0])), right: Math.max(...points.map(p => p[0])),
    top: Math.min(...points.map(p => p[1])), bottom: Math.max(...points.map(p => p[1])) });
  const exceptTextBodyAttribute = (s, owner, attribute) => {
    assert.ok(["rot", "vert"].includes(attribute));
    if (owner === "table") {
      // Prove the inherited namespace before normalizing redundant SDK
      // declarations in this one-cell fixture. Other XML stays significant.
      const graphic = s.match(/<a:graphic\b[\s\S]*?<\/a:graphic>/u)?.[0];
      assert.ok(graphic?.startsWith('<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'));
      for (const m of graphic.matchAll(/\sxmlns:a="([^"]*)"/gu)) assert.equal(m[1], "http://schemas.openxmlformats.org/drawingml/2006/main");
      const bodies = [...graphic.matchAll(/<a:txBody\b[\s\S]*?<\/a:txBody>/gu)];
      assert.equal(bodies.length, 1);
      const body = bodies[0][0], normalized = body.replace(/<a:(?:bodyPr|lstStyle|p)\b[^>]*>/gu,
        tag => tag.replace(' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"', ""));
      s = s.replace(body, normalized);
    }
    return orderedXml(s.replace(/<a:bodyPr\b[^>]*>/u, tag => tag.replace(new RegExp(`\\s${attribute}="[^"]*"`, "u"), "")));
  };
  const exceptTextRotation = (s, owner) => exceptTextBodyAttribute(s, owner, "rot");
  async function assertRotationPixels(name, receipt, input, owner, angle, reference, outer = 0, flip = false) {
    const rotation = rotationBody(receipt, owner).bodyProperties.rotation;
    assert.equal(rotation.case, angle === undefined ? undefined : "rotationAngle60000");
    assert.equal(rotation.value, angle === undefined ? undefined : angle * 60000);
    const painted = await savePaint(`text-rotation-${name}`, receipt), svg = painted.pages[0].svg;
    assert.equal(svg.includes("data-officekit-text-rotation="), angle !== undefined);
    if (angle !== undefined) assert.ok(svg.includes(`data-officekit-text-rotation="${angle}" transform="rotate(${angle} 256 224)"`));
    assert.equal(svg.includes('data-officekit-text-reflection="compensated"'), flip,
      "horizontal shape reflection must not mirror text glyphs");
    assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-rotation" ||
      d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith("rotationAngle60000")));
    assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-layout"));
    const actual = await rotationInk(painted), a = (angle ?? 0) * Math.PI / 180, b = outer * Math.PI / 180;
    // Independent point-space transform of the zero-angle glyph samples.
    // Raster hinting may differ after rotation; compare bounds, centroids and
    // bidirectional ink neighborhoods, not only SVG transform attributes.
    const expected = reference.points.map(([x, y]) => {
      const dx = x + .5 - 256, dy = y + .5 - 224;
      // Horizontal shape flip is compensated for text, unlike its outline.
      const tx = dx * Math.cos(a) - dy * Math.sin(a), ty = dx * Math.sin(a) + dy * Math.cos(a);
      return [256 + tx * Math.cos(b) - ty * Math.sin(b) - .5, 224 + tx * Math.sin(b) + ty * Math.cos(b) - .5];
    });
    const bounds = rotationBounds(actual.points), expectedBounds = rotationBounds(expected);
    for (const key of Object.keys(bounds)) assert.ok(Math.abs(bounds[key] - expectedBounds[key]) <= 2, `${name}/${key}: actual ${bounds[key]}, expected ${expectedBounds[key]}`);
    for (const axis of [0, 1]) {
      const mean = points => points.reduce((sum, p) => sum + p[axis], 0) / points.length;
      assert.ok(Math.abs(mean(actual.points) - mean(expected)) < 1.5, `${name}: glyph centroid must rotate about the text frame`);
    }
    const neighborhoods = points => new Set(points.map(([x, y]) => `${Math.round(x)},${Math.round(y)}`));
    for (const [points, set] of [[expected, neighborhoods(actual.points)], [actual.points, neighborhoods(expected)]]) for (const [x, y] of points) {
      let found = false;
      for (let dx = -2; dx <= 2 && !found; dx++) for (let dy = -2; dy <= 2 && !found; dy++) found = set.has(`${Math.round(x) + dx},${Math.round(y) + dy}`);
      assert.ok(found, `${name}: rotated glyph neighborhood missing at ${x},${y}`);
    }
    assert.ok(Math.abs(actual.points.length - reference.points.length) / reference.points.length < .2);
    if (!outer && owner !== "text") {
      const i = (162 * actual.info.width + 130) * actual.info.channels;
      assert.deepEqual([...actual.data.subarray(i, i + 4)], [245, 240, 224, 255], "text rotation must leave its fill fixed");
    }
    const entry = await assertProductionEntry(`text-rotation-${name}`, input, receipt, painted);
    return { name, owner, angle: angle ?? null, outer, flip, bounds, inkPixels: actual.points.length,
      reliability: entry.reliability.status, candidateSha256: sha256(receipt.file), independentRotatedInk: true };
  }
  for (const owner of ["text", "shape", "table"]) {
    const zero = await compileFormat(rotationProgram(owner, 0));
    const reference = await rotationInk(await savePaint(`text-rotation-${owner}-reference`, zero));
    const configurations = [-90, 0, 90, 180, 12.25].map(angle => ({ angle, outer: 0, flip: false }));
    if (owner !== "table") configurations.push({ angle: -90, outer: 90, flip: false }, { angle: 90, outer: 90, flip: true });
    for (const { angle, outer, flip } of configurations) {
      const name = `${owner}-${angle}-${outer}-${flip}`;
      try {
        const program = rotationProgram(owner, angle, outer, flip), bytes = Buffer.from(JSON.stringify(program)), inputHash = sha256(bytes);
        const authored = await compilePpjWorkspace({ program: bytes, assets: [] }, { includePreviewScene: true });
        textRotationCases.push(await assertRotationPixels(`${name}-authored`, authored, { program: bytes, assets: [] }, owner, angle, reference, outer, flip));
        const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
        await writeFile(path.join(artifacts, `text-rotation-${name}-source.pptx`), source, { flag: "wx" });
        const projected = await projectPptxToPpj(source, { sourceUri: "text-rotation.pptx", assetRootUri: "assets" });
        const input = { program: projected.programJson, source, assets: projected.assets };
        const noop = await compilePpjWorkspace(input, { includePreviewScene: true });
        assert.deepEqual(noop.file, source);
        textRotationCases.push({ ...await assertRotationPixels(`${name}-source`, noop, input, owner, angle, reference, outer, flip), sourceSha256: sourceHash });
        if (angle === 90 && !outer && !flip) for (const operation of ["negative", "zero", "delete"]) {
          try {
            const projection = await projectPptxToPpj(source, { sourceUri: "text-rotation.pptx", assetRootUri: "assets" });
            const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8")), style = rotationStyle(request, owner);
            assert.equal(style.rotation, 90);
            const value = operation === "negative" ? -90 : operation === "zero" ? 0 : undefined;
            if (value === undefined) delete style.rotation; else style.rotation = value;
            const requestBytes = Buffer.from(JSON.stringify(request)), requestHash = sha256(requestBytes), editName = `${owner}-${operation}`;
            await writeFile(path.join(artifacts, `text-rotation-${editName}.request.ppj`), requestBytes, { flag: "wx" });
            const editInput = { program: requestBytes, source, assets: projection.assets };
            const candidate = await compilePpjWorkspace(editInput, { includePreviewScene: true });
            await writeFile(path.join(artifacts, `text-rotation-${editName}.pptx`), candidate.file, { flag: "wx" });
            const result = await assertRotationPixels(editName, candidate, editInput, owner, value, reference);
            const fresh = await projectPptxToPpj(candidate.file, { sourceUri: "text-rotation-candidate.pptx", assetRootUri: "assets" });
            await writeFile(path.join(artifacts, `text-rotation-${editName}.reprojected.ppj`), fresh.programJson, { flag: "wx" });
            assert.deepEqual(rotationStyle(JSON.parse(Buffer.from(fresh.programJson).toString("utf8")), owner), style);
            assert.deepEqual(rotationBody(candidate, owner).paragraphs, rotationBody(noop, owner).paragraphs,
              "rotation edits must preserve every native paragraph, run, style and break");
            const old = await JSZip.loadAsync(source), updated = await JSZip.loadAsync(candidate.file);
            assert.deepEqual(Object.keys(updated.files).sort(), Object.keys(old.files).sort());
            const changedParts = [];
            for (const part of Object.keys(old.files)) if (!old.files[part].dir &&
              !Buffer.from(await old.file(part).async("uint8array")).equals(Buffer.from(await updated.file(part).async("uint8array")))) changedParts.push(part);
            assert.deepEqual(changedParts, ["ppt/slides/slide1.xml"]);
            const xml = await updated.file(changedParts[0]).async("string"), bodyPr = xml.match(/<a:bodyPr\b[^>]*>/)?.[0];
            assert.ok(bodyPr);
            assert.equal(/\brot=/.test(bodyPr), value !== undefined);
            if (value !== undefined) assert.ok(bodyPr.includes(`rot="${value * 60000}"`));
            const originalXml = await old.file(changedParts[0]).async("string");
            const exceptRotation = s => exceptTextRotation(s, owner);
            assert.equal(exceptRotation(xml), exceptRotation(originalXml), "only the target bodyPr rot may change inside the slide");
            assert.notEqual(exceptRotation(xml.replace(">F0<", ">ALTERED<")), exceptRotation(originalXml), "normalization must still detect changed text");
            if (owner === "table") assert.throws(() => exceptRotation(xml.replace(
              '<a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"', '<a:bodyPr xmlns:a="urn:wrong"')),
            "a changed namespace must not be normalized away");
            assert.equal(sha256(source), sourceHash); assert.equal(sha256(requestBytes), requestHash);
            textRotationCases.push({ ...result, sourceSha256: sourceHash, requestSha256: requestHash,
              reprojectionSha256: sha256(fresh.programJson), changedParts, nativeParagraphsPreserved: true, onlyBodyRotationChanged: true });
          } catch (error) { textRotationFailures.push({ name: `${owner}-${operation}`, message: error.message, code: error.code ?? null }); }
        }
        assert.equal(sha256(source), sourceHash); assert.equal(sha256(bytes), inputHash);
      } catch (error) { textRotationFailures.push({ name, message: error.message, code: error.code ?? null }); }
    }
  }
  // Retain the initially discovered input, not just its compatible explicit-
  // break counterpart. Literal newlines in a table run currently make native
  // import opaque. This proves preservation/refusal only, NOT source drawing.
  for (const angle of [0, 90]) {
    const program = rotationProgram("table", angle);
    program.pages[0].elements[0].rows[0].cells[0].text.paragraphs[0].runs = [
      { text: "F0\nIL", style: { size: 24, color: "#0055CC" } }];
    const authored = await compileFormat(program), source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    await writeFile(path.join(artifacts, `text-rotation-table-literal-newline-${angle}.pptx`), source, { flag: "wx" });
    const projection = await projectPptxToPpj(source, { sourceUri: "literal-newline.pptx", assetRootUri: "assets" });
    const p = JSON.parse(Buffer.from(projection.programJson).toString("utf8"));
    assert.equal(p.pages[0].elements[0].type, "opaque");
    const noop = await compilePpjWorkspace({ program: projection.programJson, source, assets: projection.assets }, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    assert.equal(noop.previewScene.presentation.slides[0].elements[0].content.case, "opaque");
    const painted = await savePaint(`text-rotation-table-literal-newline-${angle}`, noop);
    assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.content" && d.status === "opaque"));
    assert.match(painted.pages[0].svg, /opaque: not painted/);
    assert.doesNotMatch(painted.pages[0].svg, /data-officekit-text-rotation=/);
    assert.equal(sha256(source), sourceHash);
    textRotationOpaqueCases.push({ angle, status: "opaque", sourceSha256: sourceHash, exactNoop: true,
      sourceDrawing: "unavailable: literal newline inside table run is outside native import profile" });
  }
  const textDirectionCases = [], textDirectionFailures = [];
  const directionProgram = (owner, mode, anchor, angle = 0, outer = 0, flip = false, reference = false) => {
    const p = rotationProgram(owner, angle, outer, flip), e = p.pages[0].elements[0], style = rotationStyle(p, owner);
    Object.assign(e.frame, { rotation: outer, flipH: flip });
    style.verticalAlignment = anchor;
    if (mode !== undefined) style.verticalText = mode;
    if (reference && ["vertical", "vertical270"].includes(mode)) {
      // Independently authored horizontal equivalent: exchange frame extents
      // about fixed center (256,224) and express physical margins in reading
      // coordinates. Do not derive this reference from the rendered scene.
      Object.assign(e.frame, { x: 192, y: 96, width: 128, height: 256 });
      if (owner === "table") { e.columns[0].width = 128; e.rows[0].height = 256; }
      style.margins = mode === "vertical" ? { left: 10, top: 12, right: 14, bottom: 8 }
        : { left: 14, top: 8, right: 10, bottom: 12 };
      style.rotation = angle + (mode === "vertical" ? 90 : -90);
      delete style.verticalText;
    }
    return p;
  };
  async function assertDirectionPixels(name, receipt, input, owner, mode, reference) {
    const body = rotationBody(receipt, owner), direction = body.bodyProperties.verticalText;
    assert.equal(direction.case, mode === undefined ? undefined : "verticalTextMode");
    assert.equal(direction.value, mode);
    const painted = await savePaint(`text-direction-${name}`, receipt), svg = painted.pages[0].svg;
    assert.equal(svg.includes("data-officekit-text-direction="), mode === "vertical" || mode === "vertical270");
    if (mode === "vertical" || mode === "vertical270") assert.ok(svg.includes(`data-officekit-text-direction="${mode}"`));
    assert.ok(!painted.diagnostics.some(d => ["preview.scene.paint.text-direction", "preview.scene.paint.text-rotation"].includes(d.reason) ||
      d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith("verticalTextMode")));
    assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-layout"));
    const actual = await rotationInk(painted), bounds = rotationBounds(actual.points), expectedBounds = rotationBounds(reference.points);
    for (const key of Object.keys(bounds)) assert.ok(Math.abs(bounds[key] - expectedBounds[key]) <= 2, `${name}/${key}`);
    for (const axis of [0, 1]) {
      const mean = points => points.reduce((sum, p) => sum + p[axis], 0) / points.length;
      assert.ok(Math.abs(mean(actual.points) - mean(reference.points)) < 1.5, `${name}: direction glyph centroid`);
    }
    for (const [points, other] of [[actual.points, reference.points], [reference.points, actual.points]]) {
      const set = new Set(other.map(([x, y]) => `${x},${y}`));
      for (const [x, y] of points) {
        let found = false;
        for (let dx = -2; dx <= 2 && !found; dx++) for (let dy = -2; dy <= 2 && !found; dy++) found = set.has(`${x + dx},${y + dy}`);
        assert.ok(found, `${name}: independent horizontal-reference ink at ${x},${y}`);
      }
    }
    assert.ok(Math.abs(actual.points.length / reference.points.length - 1) < .2);
    const entry = await assertProductionEntry(`text-direction-${name}`, input, receipt, painted);
    return { name, owner, mode: mode ?? null, bounds, inkPixels: actual.points.length, independentHorizontalReference: true,
      candidateSha256: sha256(receipt.file), reliability: entry.reliability.status };
  }
  for (const owner of ["text", "shape", "table"]) {
    const profiles = ["horizontal", "vertical", "vertical270"].flatMap(mode => ["top", "middle", "bottom"].map(anchor => ({ mode, anchor, angle: 0, outer: 0, flip: false })));
    profiles.push({ mode: "vertical", anchor: "middle", angle: -90, outer: 90, flip: true },
      { mode: "vertical270", anchor: "middle", angle: 12.25, outer: 90, flip: true });
    for (const { mode, anchor, angle, outer, flip } of profiles) {
      const name = `${owner}-${mode}-${anchor}-${angle}-${outer}-${flip}`;
      try {
        const program = directionProgram(owner, mode, anchor, angle, outer, flip), input = { program: Buffer.from(JSON.stringify(program)), assets: [] };
        const inputHash = sha256(input.program), authored = await compilePpjWorkspace(input, { includePreviewScene: true });
        const referenceReceipt = await compileFormat(directionProgram(owner, mode, anchor, angle, outer, flip, true));
        const reference = await rotationInk(await savePaint(`text-direction-${name}-reference`, referenceReceipt));
        const authoredFile = `text-direction-${name}-authored-candidate.pptx`;
        await writeFile(path.join(artifacts, authoredFile), authored.file, { flag: "wx" });
        textDirectionCases.push({ ...await assertDirectionPixels(`${name}-authored`, authored, input, owner, mode, reference), candidateFile: authoredFile });
        const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source), sourceFile = `text-direction-${name}-source.pptx`;
        await writeFile(path.join(artifacts, sourceFile), source, { flag: "wx" });
        const projection = await projectPptxToPpj(source, { sourceUri: sourceFile, assetRootUri: "assets" });
        const sourceInput = { program: projection.programJson, source, assets: projection.assets };
        const noop = await compilePpjWorkspace(sourceInput, { includePreviewScene: true });
        assert.deepEqual(noop.file, source);
        textDirectionCases.push({ ...await assertDirectionPixels(`${name}-source`, noop, sourceInput, owner, mode, reference), sourceFile, sourceSha256: sourceHash, candidateFile: sourceFile });
        if (mode === "vertical" && anchor === "middle" && !angle && !outer && !flip) for (const target of ["vertical270", "horizontal", undefined]) {
          const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8")), style = rotationStyle(request, owner);
          if (target === undefined) delete style.verticalText; else style.verticalText = target;
          const editName = `${owner}-${target ?? "delete"}`, requestFile = `text-direction-${editName}.request.ppj`;
          const requestBytes = Buffer.from(JSON.stringify(request)), requestHash = sha256(requestBytes);
          await writeFile(path.join(artifacts, requestFile), requestBytes, { flag: "wx" });
          const editInput = { ...sourceInput, program: requestBytes }, candidate = await compilePpjWorkspace(editInput, { includePreviewScene: true });
          const candidateFile = `text-direction-${editName}-candidate.pptx`;
          await writeFile(path.join(artifacts, candidateFile), candidate.file, { flag: "wx" });
          const expected = await rotationInk(await savePaint(`text-direction-${editName}-reference`, await compileFormat(directionProgram(owner, target, anchor, 0, 0, false, true))));
          const record = await assertDirectionPixels(editName, candidate, editInput, owner, target, expected);
          const fresh = await projectPptxToPpj(candidate.file, { sourceUri: candidateFile, assetRootUri: "assets" });
          const reprojectionFile = `text-direction-${editName}.reprojected.ppj`;
          await writeFile(path.join(artifacts, reprojectionFile), fresh.programJson, { flag: "wx" });
          assert.deepEqual(rotationStyle(JSON.parse(Buffer.from(fresh.programJson).toString("utf8")), owner), style);
          assert.deepEqual(rotationBody(candidate, owner).paragraphs, rotationBody(noop, owner).paragraphs);
          const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file);
          assert.deepEqual(Object.keys(newZip.files).sort(), Object.keys(oldZip.files).sort());
          const changedParts = [];
          for (const part of Object.keys(oldZip.files)) if (!oldZip.files[part].dir &&
            !Buffer.from(await oldZip.file(part).async("uint8array")).equals(Buffer.from(await newZip.file(part).async("uint8array")))) changedParts.push(part);
          assert.deepEqual(changedParts, ["ppt/slides/slide1.xml"]);
          const originalXml = await oldZip.file(changedParts[0]).async("string"), xml = await newZip.file(changedParts[0]).async("string");
          const tag = xml.match(/<a:bodyPr\b[^>]*>/)?.[0]; assert.ok(tag);
          assert.equal(xmlAttributes(tag).vert, target === "horizontal" ? "horz" : target === "vertical270" ? "vert270" : undefined);
          assert.equal(exceptTextBodyAttribute(xml, owner, "vert"), exceptTextBodyAttribute(originalXml, owner, "vert"));
          assert.notEqual(exceptTextBodyAttribute(xml.replace(">F0<", ">ALTERED<"), owner, "vert"), exceptTextBodyAttribute(originalXml, owner, "vert"));
          assert.equal(sha256(requestBytes), requestHash);
          textDirectionCases.push({ ...record, sourceFile, sourceSha256: sourceHash, candidateFile, requestFile, requestSha256: requestHash,
            reprojectionFile, reprojectionSha256: sha256(fresh.programJson), onlyBodyDirectionChanged: true });
        }
        assert.equal(sha256(source), sourceHash); assert.equal(sha256(input.program), inputHash);
      } catch (error) { textDirectionFailures.push({ name, message: error.message, code: error.code ?? null }); }
    }
  }
  const textReflectionCases = [], textReflectionFailures = [];
  // Independent point-space reference, not SVG parsing or painter helpers.
  const transformPoint = ([x, y], frame) => {
    const cx = frame.x + frame.width / 2, cy = frame.y + frame.height / 2;
    const dx = (x - cx) * (frame.flipH ? -1 : 1), dy = (y - cy) * (frame.flipV ? -1 : 1);
    const a = (frame.rotation || 0) * Math.PI / 180;
    return [cx + dx * Math.cos(a) - dy * Math.sin(a), cy + dx * Math.sin(a) + dy * Math.cos(a)];
  };
  const reflectionProfiles = [
    { name: "self-h", self: { flipH: true }, groups: [] },
    { name: "self-v", self: { flipV: true }, groups: [] },
    { name: "self-both", self: { flipH: true, flipV: true }, groups: [] },
    { name: "group-h", self: {}, groups: [{ flipH: true }] },
    { name: "group-v", self: {}, groups: [{ flipV: true }] },
    { name: "nested-cancel", self: {}, groups: [{ flipH: true }, { flipH: true }] },
    { name: "nested-rotated", self: { flipH: true, rotation: 30 }, groups: [{ flipV: true, rotation: -90 }, { flipH: true, rotation: 90 }] },
    { name: "nested-scaled", self: { flipV: true }, groups: [{ flipH: true, rotation: 90 }, { flipV: true, rotation: -90, width: 320, x: 96 }] },
  ];
  for (const owner of ["shape", "table"]) {
    const reference = await rotationInk(await savePaint(`text-reflection-${owner}-reference`, await compileFormat(rotationProgram(owner, 0))));
    for (const profile of reflectionProfiles) {
      const name = `${owner}-${profile.name}`;
      try {
        const program = rotationProgram(owner, 90), leaf = program.pages[0].elements[0];
        Object.assign(leaf.frame, profile.self);
        let root = leaf;
        const groups = profile.groups.map((g, i) => ({ id: `reflection-group-${i}`, type: "group",
          frame: { ...rotationFrame, ...g }, childFrame: { ...rotationFrame } }));
        for (const group of groups) root = { ...group, elements: [root], readingOrder: [root.id] };
        program.pages[0].elements[0] = root;
        const input = { program: Buffer.from(JSON.stringify(program)), assets: [] }, inputHash = sha256(input.program);
        const authored = await compilePpjWorkspace(input, { includePreviewScene: true });
        const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
        const sourceFile = `text-reflection-${name}-source.pptx`;
        await writeFile(path.join(artifacts, sourceFile), source, { flag: "wx" });
        const projection = await projectPptxToPpj(source, { sourceUri: sourceFile, assetRootUri: "assets" });
        const sourceInput = { program: projection.programJson, source, assets: projection.assets };
        const noop = await compilePpjWorkspace(sourceInput, { includePreviewScene: true });
        assert.deepEqual(noop.file, source);
        const scenarios = [{ name: "authored", receipt: authored, input, angle: 90 }, { name: "source", receipt: noop, input: sourceInput, angle: 90 }];
        const projectedLeaf = p => {
          let e = p.pages[0].elements[0];
          for (const group of groups) { assert.equal(e.type, "group"); assert.equal(e.elements.length, 1); e = e.elements[0]; }
          assert.ok(owner === "table" ? e.type === "table" : ["shape", "text"].includes(e.type));
          return e;
        };
        const leafStyle = e => owner === "table" ? e.rows[0].cells[0].text.style : e.type === "text" ? e.style : e.textStyle;
        if (profile.name === "nested-rotated" || profile.name === "nested-scaled") {
          const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8"));
          leafStyle(projectedLeaf(request)).rotation = -90;
          const editInput = { ...sourceInput, program: Buffer.from(JSON.stringify(request)) };
          const requestFile = `text-reflection-${name}-edit.ppj`, requestHash = sha256(editInput.program);
          await writeFile(path.join(artifacts, requestFile), editInput.program, { flag: "wx" });
          const receipt = await compilePpjWorkspace(editInput, { includePreviewScene: true });
          assert.equal(sha256(editInput.program), requestHash);
          const fresh = await projectPptxToPpj(receipt.file, { sourceUri: "reflection-edited.pptx", assetRootUri: "assets" });
          assert.equal(leafStyle(projectedLeaf(JSON.parse(Buffer.from(fresh.programJson).toString("utf8")))).rotation, -90);
          const beforeZip = await JSZip.loadAsync(source), afterZip = await JSZip.loadAsync(receipt.file);
          assert.deepEqual(Object.keys(afterZip.files).sort(), Object.keys(beforeZip.files).sort());
          const changes = [];
          for (const file of Object.keys(beforeZip.files).filter(file => !beforeZip.files[file].dir)) {
            if (!Buffer.from(await beforeZip.file(file).async("uint8array")).equals(Buffer.from(await afterZip.file(file).async("uint8array")))) changes.push(file);
          }
          assert.deepEqual(changes, ["ppt/slides/slide1.xml"]);
          const oldXml = await beforeZip.file(changes[0]).async("string"), newXml = await afterZip.file(changes[0]).async("string");
          assert.equal(exceptTextRotation(newXml, owner), exceptTextRotation(oldXml, owner), "only body rotation may change, including within nested groups");
          assert.notEqual(exceptTextRotation(newXml.replace(">F0<", ">ALTERED<"), owner), exceptTextRotation(oldXml, owner));
          scenarios.push({ name: "edit", receipt, input: editInput, angle: -90, requestFile, requestHash, reprojectedSha256: sha256(fresh.programJson) });
          await writeFile(path.join(artifacts, `text-reflection-${name}-edit.reprojected.ppj`), fresh.programJson, { flag: "wx" });
        }
        const frames = [leaf.frame, ...groups.map(g => g.frame)];
        const odd = frames.reduce((v, f) => v ^ Number(!!f.flipH) ^ Number(!!f.flipV), 0) === 1;
        for (const scenario of scenarios) {
          const caseName = `${name}-${scenario.name}`, compiled = scenario.receipt;
          const candidateFile = `text-reflection-${caseName}-candidate.pptx`;
          await writeFile(path.join(artifacts, candidateFile), compiled.file, { flag: "wx" });
          let native = createPpjSceneView(compiled).pages[0].nodes[0];
          for (const group of [...groups].reverse()) {
            assert.equal(native.kind, "group");
            for (const key of ["rotation", "flipH", "flipV"]) assert.equal(native.transform?.[key] ?? (key === "rotation" ? 0 : false), group.frame[key] ?? (key === "rotation" ? 0 : false));
            assert.deepEqual(native.frame, Object.fromEntries(["x", "y", "width", "height"].map(k => [k, group.frame[k]])));
            assert.deepEqual(native.childFrame, rotationFrame);
            assert.equal(native.children.length, 1); native = native.children[0];
          }
          assert.equal(native.kind, owner === "table" ? "table" : "shape");
          for (const key of ["rotation", "flipH", "flipV"]) assert.equal(native.transform?.[key] ?? (key === "rotation" ? 0 : false), leaf.frame[key] ?? (key === "rotation" ? 0 : false));
          const body = owner === "table" ? native.native.rows[0].cells[0].textBody : native.native.textBody;
          assert.equal(body.bodyProperties.rotation.value, scenario.angle * 60000);
          if (scenario.name === "edit") {
            let originalLeaf = createPpjSceneView(noop).pages[0].nodes[0];
            for (const group of groups) originalLeaf = originalLeaf.children[0];
            const originalBody = owner === "table" ? originalLeaf.native.rows[0].cells[0].textBody : originalLeaf.native.textBody;
            assert.deepEqual(body.paragraphs, originalBody.paragraphs, "source rotation preserves every native paragraph/run/style/break");
          }
          const painted = await savePaint(`text-reflection-${caseName}`, compiled), svg = painted.pages[0].svg;
          assert.equal((svg.match(/data-officekit-text-reflection="compensated"/g) || []).length, odd ? 1 : 0);
          assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-layout"));
          const actual = await rotationInk(painted);
          const fillCenter = (224 * actual.info.width + 256) * actual.info.channels;
          assert.deepEqual([...actual.data.subarray(fillCenter, fillCenter + 4)], [245, 240, 224, 255],
            "shape/cell fill remains painted independently of reflected and rotated text");
          const expected = reference.points.map(([x, y]) => {
            let p = transformPoint([x + .5, y + .5], { ...rotationFrame, rotation: scenario.angle });
            if (odd) p = [512 - p[0], p[1]];
            p = transformPoint(p, leaf.frame);
            for (const group of groups) {
              p = [group.frame.x + (p[0] - rotationFrame.x) * group.frame.width / rotationFrame.width,
                group.frame.y + (p[1] - rotationFrame.y) * group.frame.height / rotationFrame.height];
              p = transformPoint(p, group.frame);
            }
            return [p[0] - .5, p[1] - .5];
          });
          const bounds = rotationBounds(actual.points), expectedBounds = rotationBounds(expected);
          for (const key of Object.keys(bounds)) assert.ok(Math.abs(bounds[key] - expectedBounds[key]) <= 2, `${caseName}/${key}: actual ${bounds[key]}, expected ${expectedBounds[key]}`);
          for (const axis of [0, 1]) {
            const mean = points => points.reduce((sum, p) => sum + p[axis], 0) / points.length;
            assert.ok(Math.abs(mean(actual.points) - mean(expected)) < 1.5, `${caseName}: glyph centroid`);
          }
          for (const [points, other] of [[expected, actual.points], [actual.points, expected]]) {
            const pixels = new Set(other.map(([x, y]) => `${Math.round(x)},${Math.round(y)}`));
            for (const [x, y] of points) {
              let found = false;
              for (let dx = -2; dx <= 2 && !found; dx++) for (let dy = -2; dy <= 2 && !found; dy++) found = pixels.has(`${Math.round(x) + dx},${Math.round(y) + dy}`);
              assert.ok(found, `${caseName}: unmirrored glyph neighborhood at ${x},${y}`);
            }
          }
          const areaScale = groups.reduce((scale, g) => scale * g.frame.width / rotationFrame.width * g.frame.height / rotationFrame.height, 1);
          assert.ok(Math.abs(actual.points.length / (reference.points.length * areaScale) - 1) < .2);
          const entry = await assertProductionEntry(`text-reflection-${caseName}`, scenario.input, compiled, painted);
          textReflectionCases.push({ name: caseName, sourceFile, sourceSha256: sourceHash, candidateFile,
            candidateSha256: sha256(compiled.file), requestFile: scenario.requestFile, requestSha256: scenario.requestHash,
            reprojectedSha256: scenario.reprojectedSha256, oddReflection: odd, bounds, inkPixels: actual.points.length,
            independentGlyphPixels: true, reliability: entry.reliability.status });
        }
        assert.equal(sha256(source), sourceHash); assert.equal(sha256(input.program), inputHash);
      } catch (error) { textReflectionFailures.push({ name, message: error.message, code: error.code ?? null }); }
    }
  }
  const { assessPpjPreviewInput } = await import("../src/ppj/preview-input-assessment.mjs");
  const hiddenProgram = structuredClone(pairBase);
  hiddenProgram.pages[0].elements = [{ id: "hidden-probe", type: "shape", hidden: true,
    frame: { x: 100, y: 100, width: 60, height: 60, rotation: 30 },
    geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } }];
  const hiddenAuthored = await compileFormat(hiddenProgram);
  const visibleProgram = structuredClone(hiddenProgram);
  visibleProgram.pages[0].elements[0].hidden = false;
  const visiblePaint = await savePaint("visibility-visible", await compileFormat(visibleProgram));
  async function visibilityPixel(painted) {
    return [...await sharp(Buffer.from(painted.pages[0].svg)).extract({ left: 130, top: 130, width: 1, height: 1 })
      .removeAlpha().raw().toBuffer()];
  }
  assert.deepEqual(await visibilityPixel(visiblePaint), [204, 85, 0]);
  const hiddenSource = await withoutAuthoredSnapshot(hiddenAuthored.file);
  const hiddenProjection = await projectPptxToPpj(hiddenSource, { sourceUri: "hidden.pptx", assetRootUri: "assets" });
  const hiddenNoop = await compilePpjWorkspace({ program: hiddenProjection.programJson, source: hiddenSource, assets: hiddenProjection.assets }, { includePreviewScene: true });
  for (const [name, compiled] of [["authored", hiddenAuthored], ["source", hiddenNoop]]) {
    const painted = await savePaint(`visibility-${name}`, compiled);
    const program = JSON.parse(new TextDecoder().decode(compiled.programJson));
    assert.equal(program.pages[0].elements[0].hidden, true);
    const legacy = assessPpjPreviewInput(program);
    // Resolve the registry reason instead of assuming a parallel vocabulary.
    const registry = JSON.parse(await readFile(new URL("../src/ppj/capability-registry.json", import.meta.url)));
    const reason = registry.previewSupport.factual.visibility.reason;
    assert.ok(legacy.diagnostics.some(d => d.reason === reason));
    const options = { rendererProfile: "native-scene-svg", sceneReceipt: compiled, scenePaint: painted };
    const assessed = assessPpjPreviewInput(program, options);
    assert.ok(!assessed.diagnostics.some(d => d.reason === reason));
    assert.ok(painted.hiddenScenePaths.length);
    assert.match(painted.pages[0].svg, /data-officekit-scene-path="[^"]+"[^>]*display="none"/);
    assert.deepEqual(await visibilityPixel(painted), [255, 255, 255]);
    const uncaptured = assessPpjPreviewInput(program, { ...options, scenePaint: { ...painted, hiddenScenePaths: [] } });
    assert.ok(uncaptured.diagnostics.some(d => d.reason === reason));
    assert.throws(() => assessPpjPreviewInput(program, { rendererProfile: "native-scene-svg" }));
    assert.throws(() => assessPpjPreviewInput({ ...program, id: "wrong-input" }, options));
    assert.throws(() => assessPpjPreviewInput(program, { ...options, scenePaint: { ...painted, scene: {} } }));
    assert.throws(() => assessPpjPreviewInput(program, { ...options, registry: { ...registry, previewScene: {} } }));
    assert.throws(() => assessPpjPreviewInput(program, { rendererProfile: "unknown" }));
    for (const d of legacy.diagnostics.filter(d => d.reason !== reason))
      assert.ok(assessed.diagnostics.some(a => a.reason === d.reason && a.path === d.path), "unrelated limits remain");
  }
  const transformProfileCases = [];
  const transformProfileFailures = [];
  const transformSourceEdits = [];
  for (const [name, transform, expected] of [
    ["rotate", { rotation: 90 }, [180, 120]],
    ["flip-h", { flipH: true }, [180, 120]],
    ["flip-v", { flipV: true }, [120, 180]],
    ["combined", { rotation: 90, flipH: true, flipV: true }, [120, 180]],
  ]) {
    const authoredProgram = structuredClone(pairBase);
    authoredProgram.pages[0].elements = [{ id: "transformed-group", type: "group",
      frame: { x: 100, y: 100, width: 100, height: 100, ...transform },
      childFrame: { x: 0, y: 0, width: 100, height: 100 }, elements: [{ id: "corner", type: "shape",
        frame: { x: 10, y: 10, width: 20, height: 20 }, geometry: { kind: "preset", preset: "rect" },
        style: { fill: { type: "solid", color: "#CC5500" } } }] }];
    await writeFile(path.join(artifacts, `profile-${name}.ppj`), JSON.stringify(authoredProgram, null, 2), { flag: "wx" });
    console.log(`Transform profile ${name}: authored compile`);
    const authored = await compileFormat(authoredProgram);
    await writeFile(path.join(artifacts, `profile-${name}.pptx`), authored.file, { flag: "wx" });
    const source = await withoutAuthoredSnapshot(authored.file);
    const sourceBefore = source.slice();
    await writeFile(path.join(artifacts, `profile-${name}-source.pptx`), source, { flag: "wx" });
    const inputs = [["authored", authored]];
    let stage = "source projection";
    try {
      const projection = await projectPptxToPpj(source, { sourceUri: "transformed.pptx", assetRootUri: "assets" });
      const projectedGroup = JSON.parse(new TextDecoder().decode(projection.programJson)).pages[0].elements[0];
      assert.deepEqual(projectedGroup.readingOrder, projectedGroup.elements.map(element => element.id));
      stage = "source no-op";
      const sourceResult = await compilePpjWorkspace({ program: projection.programJson, source, assets: projection.assets }, { includePreviewScene: true });
      assert.deepEqual(sourceResult.file, source);
      inputs.push(["source", sourceResult]);
      // Rebuild every edit from the original projection and exact source, not
      // a prior candidate whose binding or transform presence may differ.
      for (const edit of ["reset", "delete"]) {
        const evidence = { sourceSha256: sha256(sourceBefore) };
        try {
          const requested = JSON.parse(new TextDecoder().decode(projection.programJson));
          const frame = requested.pages[0].elements[0].frame;
          for (const key of Object.keys(transform)) {
            if (edit === "delete") delete frame[key];
            else frame[key] = key === "rotation" ? 0 : false;
          }
          const candidate = await compilePpjWorkspace({ program: Buffer.from(JSON.stringify(requested)), source, assets: projection.assets }, { includePreviewScene: true });
          evidence.candidateSha256 = sha256(candidate.file);
          evidence.sceneSha256 = candidate.previewScene.sha256;
          await writeFile(path.join(artifacts, `profile-${name}-${edit}.pptx`), candidate.file, { flag: "wx" });
          const painted = await savePaint(`profile-${name}-${edit}`, candidate);
          const reprojection = await projectPptxToPpj(candidate.file, { sourceUri: "candidate.pptx", assetRootUri: "assets" });
          await writeFile(path.join(artifacts, `profile-${name}-${edit}.reprojected.ppj`), reprojection.programJson, { flag: "wx" });
          const freshGroup = JSON.parse(new TextDecoder().decode(reprojection.programJson)).pages[0].elements[0];
          evidence.requestedFrame = frame;
          evidence.observedFrame = freshGroup.frame;
          const { data, info } = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
          const pixel = (x, y) => [...data.subarray((y * info.width + x) * info.channels, (y * info.width + x) * info.channels + 3)];
          assert.deepEqual(pixel(120, 120), [204, 85, 0], `${name}/${edit}: reset child position`);
          assert.deepEqual(pixel(...expected), [255, 255, 255], `${name}/${edit}: old position is empty`);
          const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file), changedParts = [];
          assert.deepEqual(Object.keys(newZip.files).sort(), Object.keys(oldZip.files).sort());
          for (const member of Object.keys(oldZip.files)) if (!oldZip.files[member].dir &&
            !Buffer.from(await oldZip.file(member).async("uint8array")).equals(Buffer.from(await newZip.file(member).async("uint8array")))) changedParts.push(member);
          assert.deepEqual(changedParts, ["ppt/slides/slide1.xml"]);
          assert.deepEqual(freshGroup.readingOrder, projectedGroup.readingOrder);
          assert.deepEqual(freshGroup.elements.map(e => e.frame), projectedGroup.elements.map(e => e.frame));
          for (const key of Object.keys(transform)) {
            if (edit === "delete") assert.ok(!Object.hasOwn(freshGroup.frame, key), `${name}/${edit}: ${key} must be absent, not explicit zero/false`);
            else assert.equal(freshGroup.frame[key], key === "rotation" ? 0 : false, `${name}/${edit}: explicit transform value preserved`);
          }
          assert.deepEqual(source, sourceBefore);
          transformSourceEdits.push({ name, edit, pixels: true, reprojection: true, changedParts,
            sourceSha256: sha256(source), candidateSha256: sha256(candidate.file), sceneSha256: candidate.previewScene.sha256 });
        } catch (error) {
          transformProfileFailures.push({ name, stage: `source ${edit}`, code: error.code, message: error.message, ...evidence });
          console.error(`Transform profile ${name} failed at source ${edit}: ${error.message}`);
        }
      }
    } catch (error) {
      transformProfileFailures.push({ name, stage, code: error.code, message: error.message });
      console.error(`Transform profile ${name} failed at ${stage}: ${error.message}`);
    }
    for (const [origin, compiled] of inputs) {
      const painted = await savePaint(`profile-${name}-${origin}`, compiled);
      const program = JSON.parse(new TextDecoder().decode(compiled.programJson));
      const registry = JSON.parse(await readFile(new URL("../src/ppj/capability-registry.json", import.meta.url)));
      const reason = registry.previewSupport.factual.transform.reason;
      const legacy = assessPpjPreviewInput(program);
      const options = { rendererProfile: "native-scene-svg", sceneReceipt: compiled, scenePaint: painted };
      const assessed = assessPpjPreviewInput(program, options);
      for (const key of Object.keys(transform)) {
        const at = `$.pages[0].elements[0].frame.${key}`;
        assert.ok(legacy.diagnostics.some(d => d.reason === reason && d.path === at));
        assert.ok(!assessed.diagnostics.some(d => d.reason === reason && d.path === at), `${name}/${origin}/${key}`);
      }
      const uncaptured = assessPpjPreviewInput(program, { ...options, scenePaint: { ...painted, transformedScenePaths: [] } });
      assert.ok(uncaptured.diagnostics.some(d => d.reason === reason));
      const groupReason = registry.previewSupport.factual.groupCoordinates.reason;
      assert.ok(legacy.diagnostics.some(d => d.reason === groupReason));
      assert.ok(!assessed.diagnostics.some(d => d.reason === groupReason));
      assert.ok(uncaptured.diagnostics.some(d => d.reason === groupReason));
      for (const d of legacy.diagnostics.filter(d => d.reason !== reason && d.reason !== groupReason))
        assert.ok(assessed.diagnostics.some(a => a.reason === d.reason && a.path === d.path));
      const { data, info } = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
      const pixel = (x, y) => [...data.subarray((y * info.width + x) * info.channels, (y * info.width + x) * info.channels + 3)];
      assert.deepEqual(pixel(...expected), [204, 85, 0]);
      assert.deepEqual(pixel(120, 120), [255, 255, 255]);
      transformProfileCases.push({ name, origin, expected, pixels: true, otherRulesRetained: true,
        sourceSha256: origin === "source" ? sha256(sourceBefore) : null,
        candidateSha256: sha256(compiled.file), sceneSha256: compiled.previewScene.sha256,
        sourceNoop: origin === "source", readingOrderVerified: origin === "source" });
    }
    assert.deepEqual(source, sourceBefore, `${name}: projection and preview preserve original source bytes`);
  }
  const groupCoordinateCases = [];
  for (const [name, childWidth, expected] of [["nested-scaled", 100, [150, 136]], ["nested-wide-child", 200, [125, 136]]]) {
    const requested = structuredClone(pairBase);
    requested.pages[0].elements = [{ id: "outer", type: "group", frame: { x: 100, y: 100, width: 200, height: 100 },
      childFrame: { x: 10, y: 20, width: childWidth, height: 50 }, elements: [{ id: "inner", type: "group",
        frame: { x: 20, y: 30, width: 40, height: 20 }, childFrame: { x: 0, y: 0, width: 20, height: 10 },
        elements: [{ id: "corner", type: "shape", frame: { x: 5, y: 2, width: 5, height: 4 },
          geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } }] }] }];
    const authored = await compileFormat(requested);
    const originalSource = await withoutAuthoredSnapshot(authored.file), beforeSource = originalSource.slice();
    await writeFile(path.join(artifacts, `${name}-source.pptx`), originalSource, { flag: "wx" });
    const projection = await projectPptxToPpj(originalSource, { sourceUri: "nested.pptx", assetRootUri: "assets" });
    const noop = await compilePpjWorkspace({ program: projection.programJson, source: originalSource, assets: projection.assets }, { includePreviewScene: true });
    assert.deepEqual(noop.file, originalSource);
    for (const [origin, compiled] of [["authored", authored], ["source", noop]]) {
      const painted = await savePaint(`${name}-${origin}`, compiled);
      const program = JSON.parse(new TextDecoder().decode(compiled.programJson));
      const registry = JSON.parse(await readFile(new URL("../src/ppj/capability-registry.json", import.meta.url)));
      const reason = registry.previewSupport.factual.groupCoordinates.reason;
      const legacy = assessPpjPreviewInput(program);
      const options = { rendererProfile: "native-scene-svg", sceneReceipt: compiled, scenePaint: painted };
      const assessed = assessPpjPreviewInput(program, options);
      assert.equal(legacy.diagnostics.filter(d => d.reason === reason).length, 2);
      assert.ok(!assessed.diagnostics.some(d => d.reason === reason));
      const uncaptured = assessPpjPreviewInput(program, { ...options, scenePaint: { ...painted, transformedScenePaths: [] } });
      assert.equal(uncaptured.diagnostics.filter(d => d.reason === reason).length, 2);
      const withoutMapping = structuredClone(registry);
      delete withoutMapping.previewScene.factualMappings.groupCoordinates;
      assert.throws(() => assessPpjPreviewInput(program, { ...options, registry: withoutMapping }), /group coordinate mapping/);
      for (const d of legacy.diagnostics.filter(d => d.reason !== reason))
        assert.ok(assessed.diagnostics.some(a => a.reason === d.reason && a.path === d.path));
      assert.ok(!paintPpjSceneSvg(compiled, { assessInput: true }).diagnostics.some(d => d.reason === reason));
      const { data, info } = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
      const pixel = (x, y) => [...data.subarray((y * info.width + x) * info.channels, (y * info.width + x) * info.channels + 3)];
      assert.deepEqual(pixel(...expected), [204, 85, 0]);
      assert.deepEqual(pixel(name === "nested-scaled" ? 125 : 150, 136), [255, 255, 255]);
      groupCoordinateCases.push({ name, origin, expected, pixels: true, twoOwners: true, missingCaptureRetained: true,
        sourceSha256: origin === "source" ? sha256(originalSource) : null, candidateSha256: sha256(compiled.file) });
    }
    assert.deepEqual(originalSource, beforeSource);
  }
  async function formatPixels(name, result, baseline, decorated, spacing = 0) {
    const painted = await savePaint(name, result);
    const run = createPpjSceneView(result).pages[0].nodes.find(n => n.kind === "shape").native.textBody.paragraphs[0].runs[0];
    assert.equal(run.fontBaselinePercent, baseline);
    assert.equal(run.underline, decorated ? "sng" : "none");
    assert.equal(run.strike, decorated ? "sngStrike" : "noStrike");
    assert.equal(run.fontSpacingPoints, spacing);
    assert.ok(painted.pages[0].svg.includes(`letter-spacing="${spacing}"`));
    assert.ok(painted.pages[0].svg.includes(`text-decoration="${decorated ? "underline line-through" : "none"}" dy="${baseline ? -40 * baseline / 100 : 0}"`));
    const { data, info } = await sharp(Buffer.from(painted.pages[0].svg)).flatten({ background: "#FFFFFF" }).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    let minY = Infinity, maxY = -Infinity, minX = Infinity, maxX = -Infinity, count = 0;
    for (let y = 60; y < 200; y++) for (let x = 100; x < 350; x++) {
      const i = (y * info.width + x) * info.channels;
      if (data[i] < 128 && data[i + 1] < 128 && data[i + 2] < 128) {
        minY = Math.min(minY, y); maxY = Math.max(maxY, y);
        minX = Math.min(minX, x); maxX = Math.max(maxX, x); count++;
      }
    }
    assert.ok(count > 0);
    return { minY, maxY, minX, maxX, count };
  }
  const plainFormat = await compileFormat(formatProgram);
  const plainPixels = await formatPixels("text-plain", plainFormat, 0, false);
  for (const spacing of [-3, 4]) {
    const spaced = structuredClone(formatProgram);
    spaced.pages[0].elements[0].text.paragraphs[0].runs[0].style.letterSpacing = spacing;
    const pixels = await formatPixels(`text-spacing-${spacing}`, await compileFormat(spaced), 0, false, spacing);
    // Four identical glyphs have three internal character gaps. Compare ink
    // width, not the backend's trailing advance or font-specific glyph size.
    assert.equal(pixels.maxX - pixels.minX, plainPixels.maxX - plainPixels.minX + 3 * spacing);
    assert.equal(pixels.minY, plainPixels.minY);
    assert.equal(pixels.maxY, plainPixels.maxY);
  }
  const shiftedProgram = structuredClone(formatProgram);
  shiftedProgram.pages[0].elements[0].text.paragraphs[0].runs[0].style.baseline = 30;
  const shiftedPixels = await formatPixels("text-raised", await compileFormat(shiftedProgram), 30, false);
  assert.equal(shiftedPixels.minY, plainPixels.minY - 12);
  assert.equal(shiftedPixels.maxY, plainPixels.maxY - 12);
  const decoratedProgram = structuredClone(formatProgram);
  Object.assign(decoratedProgram.pages[0].elements[0].text.paragraphs[0].runs[0].style, { underline: "sng", strike: "sngStrike" });
  const decoratedPixels = await formatPixels("text-decorated", await compileFormat(decoratedProgram), 0, true);
  assert.ok(decoratedPixels.count > plainPixels.count, "Decoration adds actual raster ink, not only SVG attributes");
  const formatSource = await withoutAuthoredSnapshot(plainFormat.file), formatBefore = formatSource.slice();
  const formatProjection = await projectPptxToPpj(formatSource, { sourceUri: "text-format.pptx", assetRootUri: "assets" });
  const formatInput = { program: formatProjection.programJson, source: formatSource, assets: formatProjection.assets };
  const formatNoop = await compilePpjWorkspace(formatInput, { includePreviewScene: true });
  assert.deepEqual(formatNoop.file, formatSource);
  await formatPixels("text-source-noop", formatNoop, 0, false);
  const formatEdit = JSON.parse(new TextDecoder().decode(formatProjection.programJson));
  for (const [kind, value] of [["fontBaselinePercent", -30], ["fontUnderline", "sng"], ["fontStrike", "sngStrike"], ["fontSpacingPoints", -3]]) {
    const leaf = formatEdit.pages[0].elements[0].nativeRef.leaves.find(leaf => leaf.kind === kind);
    assert.ok(leaf, `Source must issue ${kind} editing capability`);
    leaf.value = value;
  }
  const formatCandidate = await compilePpjWorkspace({ ...formatInput, program: Buffer.from(JSON.stringify(formatEdit)) }, { includePreviewScene: true });
  const editedPixels = await formatPixels("text-source-edited", formatCandidate, -30, true, -3);
  assert.equal(editedPixels.minY, decoratedPixels.minY + 12);
  assert.equal(editedPixels.maxY, decoratedPixels.maxY + 12);
  const equivalentFormat = structuredClone(decoratedProgram);
  Object.assign(equivalentFormat.pages[0].elements[0].text.paragraphs[0].runs[0].style, { baseline: -30, letterSpacing: -3 });
  assert.deepEqual(editedPixels, await formatPixels("text-equivalent-authored", await compileFormat(equivalentFormat), -30, true, -3));
  const formatFresh = await projectPptxToPpj(formatCandidate.file, { sourceUri: "text-format-edited.pptx", assetRootUri: "assets" });
  const freshStyle = JSON.parse(new TextDecoder().decode(formatFresh.programJson)).pages[0].elements[0].text.paragraphs[0].runs[0].style;
  assert.equal(freshStyle.baseline, -30);
  assert.equal(freshStyle.underline, "single");
  assert.equal(freshStyle.strike, "sngStrike");
  assert.equal(freshStyle.letterSpacing, -3);
  const formatOldZip = await JSZip.loadAsync(formatSource), formatNewZip = await JSZip.loadAsync(formatCandidate.file);
  assert.deepEqual(Object.keys(formatOldZip.files).sort(), Object.keys(formatNewZip.files).sort());
  const formatChangedParts = [];
  for (const name of Object.keys(formatOldZip.files)) if (!formatOldZip.files[name].dir &&
    !Buffer.from(await formatOldZip.file(name).async("uint8array")).equals(Buffer.from(await formatNewZip.file(name).async("uint8array")))) formatChangedParts.push(name);
  assert.deepEqual(formatChangedParts, ["ppt/slides/slide1.xml"]);
  assert.deepEqual(formatSource, formatBefore);
  const paragraphProgram = structuredClone(formatProgram);
  paragraphProgram.pages[0].elements[0].frame.height = 260;
  paragraphProgram.pages[0].elements[0].text.paragraphs = [
    { style: { lineSpacing: 40, spaceBefore: 10, spaceAfter: 15, indent: 30, hanging: 10 }, runs: [{ text: "HH\nHH", style: { size: 20, color: "#000000" } }] },
    { runs: [{ text: "HH", style: { size: 20, color: "#000000" } }] },
  ];
  async function paragraphPixels(name, result) {
    const painted = await savePaint(name, result);
    const positions = [...painted.pages[0].svg.matchAll(/<text x="([^"]+)" y="([^"]+)" text-anchor/g)].map(m => ({ x: Number(m[1]), y: Number(m[2]) }));
    const ys = positions.map(p => p.y), xs = positions.map(p => p.x);
    assert.equal(ys.length, 3);
    const { data, info } = await sharp(Buffer.from(painted.pages[0].svg)).flatten({ background: "#FFFFFF" }).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const starts = [], lefts = []; let previousInk = false;
    for (let y = 60; y < 350; y++) {
      let ink = false, left = Infinity;
      for (let x = 100; x < 250; x++) {
        const i = (y * info.width + x) * info.channels;
        if (data[i] < 128 && data[i + 1] < 128 && data[i + 2] < 128) { ink = true; left = x; break; }
      }
      if (ink && !previousInk) { starts.push(y); lefts.push(left); }
      else if (ink) lefts[lefts.length - 1] = Math.min(lefts.at(-1), left);
      previousInk = ink;
    }
    assert.equal(starts.length, 3);
    return { ys, xs, starts, lefts };
  }
  const multiplierCases = [];
  for (const multiplier of [1, 1.5, 2]) {
    const program = structuredClone(paragraphProgram);
    program.pages[0].elements[0].text.paragraphs = [{ style: { lineSpacingMultiplier: multiplier },
      runs: [{ text: "HH\nHH\nHH", style: { size: 20, color: "#000000" } }] }];
    const originalProgram = JSON.stringify(program);
    const equivalent = structuredClone(program);
    equivalent.pages[0].elements[0].text.paragraphs[0].style = { lineSpacing: 24 * multiplier };
    const expected = await paragraphPixels(`multiplier-${multiplier}-point-reference`, await compileFormat(equivalent));
    const authored = await compileFormat(program);
    const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    const projected = await projectPptxToPpj(source, { sourceUri: "multiplier.pptx", assetRootUri: "assets" });
    const input = { program: projected.programJson, source, assets: projected.assets };
    const noop = await compilePpjWorkspace(input, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    for (const [origin, receipt] of [["authored", authored], ["source", noop]]) {
      const p = receipt.previewScene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0];
      assert.deepEqual(p.lineSpacing, { case: "lineSpacingMultiplier", value: multiplier });
      const observed = await paragraphPixels(`multiplier-${multiplier}-${origin}`, receipt);
      assert.deepEqual(observed, expected);
      assert.equal(observed.starts[1] - observed.starts[0], 24 * multiplier);
      assert.equal(observed.starts[2] - observed.starts[1], 24 * multiplier);
      multiplierCases.push({ multiplier, origin, pixels: observed, candidateSha256: sha256(receipt.file) });
    }
    if (multiplier === 1) {
      const edited = JSON.parse(new TextDecoder().decode(projected.programJson));
      const leaf = edited.pages[0].elements[0].nativeRef.leaves.find(item => item.kind === "paragraphLineSpacingMultiplier");
      assert.ok(leaf, "Source must issue multiplier editing capability");
      leaf.value = 2;
      const candidate = await compilePpjWorkspace({ ...input, program: Buffer.from(JSON.stringify(edited)) }, { includePreviewScene: true });
      equivalent.pages[0].elements[0].text.paragraphs[0].style.lineSpacing = 48;
      assert.deepEqual(await paragraphPixels("multiplier-source-edited", candidate),
        await paragraphPixels("multiplier-edited-point-reference", await compileFormat(equivalent)));
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: "multiplier-edited.pptx", assetRootUri: "assets" });
      assert.equal(JSON.parse(new TextDecoder().decode(fresh.programJson)).pages[0].elements[0].text.paragraphs[0].style.lineSpacingMultiplier, 2);
      const before = await JSZip.loadAsync(source), after = await JSZip.loadAsync(candidate.file), changed = [];
      assert.deepEqual(Object.keys(after.files).sort(), Object.keys(before.files).sort());
      for (const name of Object.keys(before.files)) if (!before.files[name].dir &&
        !Buffer.from(await before.file(name).async("uint8array")).equals(Buffer.from(await after.file(name).async("uint8array")))) changed.push(name);
      assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
      multiplierCases.push({ multiplier: 2, origin: "source-edit", reprojection: true, changedParts: changed, candidateSha256: sha256(candidate.file) });
    }
    assert.equal(sha256(source), sourceHash);
    assert.equal(JSON.stringify(program), originalProgram);
  }
  const paragraphMultiplierCases = [];
  for (const size of [10, 20]) for (const multiplier of [0, .5]) {
    const program = structuredClone(paragraphProgram);
    program.pages[0].elements[0].text.paragraphs = Array.from({ length: 3 }, () => ({
      style: { spaceBeforeMultiplier: multiplier, spaceAfterMultiplier: multiplier },
      runs: [{ text: "HH", style: { size, color: "#000000" } }],
    }));
    const originalProgram = JSON.stringify(program), equivalent = structuredClone(program);
    for (const p of equivalent.pages[0].elements[0].text.paragraphs)
      p.style = { spaceBefore: size * 1.2 * multiplier, spaceAfter: size * 1.2 * multiplier };
    const expected = await paragraphPixels(`paragraph-multiplier-${size}-${multiplier}-reference`, await compileFormat(equivalent));
    const authored = await compileFormat(program), source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    const projection = await projectPptxToPpj(source, { sourceUri: "paragraph-multiplier.pptx", assetRootUri: "assets" });
    const input = { program: projection.programJson, source, assets: projection.assets };
    const noop = await compilePpjWorkspace(input, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    for (const [origin, receipt] of [["authored", authored], ["source", noop]]) {
      const paragraphs = receipt.previewScene.presentation.slides[0].elements[0].content.value.textBody.paragraphs;
      for (const p of paragraphs) for (const field of ["spaceBefore", "spaceAfter"])
        assert.deepEqual(p[field], { case: `${field}Multiplier`, value: multiplier });
      const pixels = await paragraphPixels(`paragraph-multiplier-${size}-${multiplier}-${origin}`, receipt);
      assert.deepEqual(pixels, expected);
      assert.equal(pixels.starts[1] - pixels.starts[0], size * 1.2 * (1 + 2 * multiplier));
      paragraphMultiplierCases.push({ size, multiplier, origin, pixels, candidateSha256: sha256(receipt.file) });
    }
    if (multiplier === .5) {
      const edited = JSON.parse(new TextDecoder().decode(projection.programJson));
      for (const kind of ["paragraphSpaceBeforeMultiplier", "paragraphSpaceAfterMultiplier"]) {
        const leaves = edited.pages[0].elements[0].nativeRef.leaves.filter(leaf => leaf.kind === kind);
        assert.equal(leaves.length, 3);
        for (const leaf of leaves) leaf.value = 0;
      }
      const candidate = await compilePpjWorkspace({ ...input, program: Buffer.from(JSON.stringify(edited)) }, { includePreviewScene: true });
      for (const p of equivalent.pages[0].elements[0].text.paragraphs) p.style = { spaceBefore: 0, spaceAfter: 0 };
      assert.deepEqual(await paragraphPixels(`paragraph-multiplier-${size}-source-zero`, candidate),
        await paragraphPixels(`paragraph-multiplier-${size}-zero-reference`, await compileFormat(equivalent)));
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: "paragraph-multiplier-zero.pptx", assetRootUri: "assets" });
      for (const p of JSON.parse(new TextDecoder().decode(fresh.programJson)).pages[0].elements[0].text.paragraphs) {
        assert.equal(p.style.spaceBeforeMultiplier, 0);
        assert.equal(p.style.spaceAfterMultiplier, 0);
      }
      const before = await JSZip.loadAsync(source), after = await JSZip.loadAsync(candidate.file), changed = [];
      assert.deepEqual(Object.keys(after.files).sort(), Object.keys(before.files).sort());
      for (const name of Object.keys(before.files)) if (!before.files[name].dir &&
        !Buffer.from(await before.file(name).async("uint8array")).equals(Buffer.from(await after.file(name).async("uint8array")))) changed.push(name);
      assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
      paragraphMultiplierCases.push({ size, multiplier: 0, origin: "source-edit", reprojection: true, changedParts: changed });
    }
    assert.equal(sha256(source), sourceHash);
    assert.equal(JSON.stringify(program), originalProgram);
  }
  const spacingDeletionCases = [], spacingDeletionFailures = [];
  const spacingDeletionProgram = structuredClone(paragraphProgram);
  spacingDeletionProgram.pages[0].elements[0].text.paragraphs[0].style = {
    lineSpacingMultiplier: 1.5, spaceBeforeMultiplier: .5, spaceAfterMultiplier: .5,
  };
  const spacingDeletionAuthored = await compileFormat(spacingDeletionProgram);
  const spacingDeletionSource = await withoutAuthoredSnapshot(spacingDeletionAuthored.file);
  const spacingDeletionHash = sha256(spacingDeletionSource), spacingDeletionZip = await JSZip.loadAsync(spacingDeletionSource);
  await writeFile(path.join(artifacts, "spacing-delete-source.pptx"), spacingDeletionSource, { flag: "wx" });
  const spacingFields = [["lineSpacingMultiplier", "lineSpacing", "lnSpc", 1.5],
    ["spaceBeforeMultiplier", "spaceBefore", "spcBef", .5], ["spaceAfterMultiplier", "spaceAfter", "spcAft", .5]];
  for (let mask = 1; mask < 8; mask++) {
    try {
      const projected = await projectPptxToPpj(spacingDeletionSource, { sourceUri: "spacing-delete.pptx", assetRootUri: "assets" });
      const request = JSON.parse(Buffer.from(projected.programJson).toString("utf8")), expected = structuredClone(spacingDeletionProgram);
      const style = request.pages[0].elements[0].text.paragraphs[0].style;
      for (const [index, [field, , , value]] of spacingFields.entries()) {
        assert.equal(style[field], value);
        if (mask & (1 << index)) {
          delete style[field];
          delete expected.pages[0].elements[0].text.paragraphs[0].style[field];
        }
      }
      const requestBytes = Buffer.from(JSON.stringify(request)), requestHash = sha256(requestBytes);
      await writeFile(path.join(artifacts, `spacing-delete-${mask}.request.ppj`), requestBytes, { flag: "wx" });
      const candidate = await compilePpjWorkspace({ program: requestBytes, source: spacingDeletionSource, assets: projected.assets }, { includePreviewScene: true });
      await writeFile(path.join(artifacts, `spacing-delete-${mask}.pptx`), candidate.file, { flag: "wx" });
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: "spacing-deleted.pptx", assetRootUri: "assets" });
      await writeFile(path.join(artifacts, `spacing-delete-${mask}.reprojected.ppj`), fresh.programJson, { flag: "wx" });
      const actualStyle = JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0].elements[0].text.paragraphs[0].style ?? {};
      const native = candidate.previewScene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0];
      const zip = await JSZip.loadAsync(candidate.file), xml = await zip.file("ppt/slides/slide1.xml").async("string");
      for (const [index, [field, slot, tag, value]] of spacingFields.entries()) {
        const deleted = Boolean(mask & (1 << index));
        assert.equal(Object.hasOwn(actualStyle, field), !deleted, `${field}: deletion must be absent, not an explicit default`);
        assert.equal(actualStyle[field], deleted ? undefined : value);
        assert.equal(native[slot].case, deleted ? undefined : field);
        assert.equal(native[slot].value, deleted ? undefined : value);
        assert.equal(new RegExp(`<a:${tag}\\b`).test(xml), !deleted, `${tag}: XML presence must match request`);
      }
      assert.deepEqual(Object.keys(zip.files).sort(), Object.keys(spacingDeletionZip.files).sort());
      const changed = [];
      for (const name of Object.keys(spacingDeletionZip.files)) if (!spacingDeletionZip.files[name].dir &&
        !Buffer.from(await spacingDeletionZip.file(name).async("uint8array")).equals(Buffer.from(await zip.file(name).async("uint8array")))) changed.push(name);
      assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
      const pixels = await paragraphPixels(`spacing-delete-${mask}`, candidate);
      assert.deepEqual(pixels, await paragraphPixels(`spacing-delete-${mask}-reference`, await compileFormat(expected)));
      assert.equal(sha256(requestBytes), requestHash);
      assert.equal(sha256(spacingDeletionSource), spacingDeletionHash);
      spacingDeletionCases.push({ mask, actualStyle, pixels, changedParts: changed, candidateSha256: sha256(candidate.file) });
    } catch (error) {
      spacingDeletionFailures.push({ mask, message: error.message, code: error.code ?? null });
      console.error(`Spacing deletion ${mask} failed: ${error.message}`);
    }
  }
  assert.equal(sha256(spacingDeletionSource), spacingDeletionHash);
  const paragraphAuthored = await compileFormat(paragraphProgram);
  const paragraphPixelsBefore = await paragraphPixels("paragraph-authored", paragraphAuthored);
  assert.ok(Math.abs(paragraphPixelsBefore.ys[1] - paragraphPixelsBefore.ys[0] - 40) < 1e-9);
  assert.equal(paragraphPixelsBefore.starts[1] - paragraphPixelsBefore.starts[0], 40);
  paragraphPixelsBefore.xs.forEach((x, i) => assert.ok(Math.abs(x - [127.2, 137.2, 107.2][i]) < 1e-9));
  assert.equal(paragraphPixelsBefore.lefts[1] - paragraphPixelsBefore.lefts[0], 10);
  const paragraphSource = await withoutAuthoredSnapshot(paragraphAuthored.file), paragraphBefore = paragraphSource.slice();
  const paragraphProjection = await projectPptxToPpj(paragraphSource, { sourceUri: "paragraph.pptx", assetRootUri: "assets" });
  const paragraphInput = { program: paragraphProjection.programJson, source: paragraphSource, assets: paragraphProjection.assets };
  const paragraphNoop = await compilePpjWorkspace(paragraphInput, { includePreviewScene: true });
  assert.deepEqual(paragraphNoop.file, paragraphSource);
  assert.deepEqual(await paragraphPixels("paragraph-noop", paragraphNoop), paragraphPixelsBefore);
  const paragraphEdit = JSON.parse(new TextDecoder().decode(paragraphProjection.programJson));
  for (const [kind, value] of [["paragraphLineSpacingPoints", 55], ["paragraphSpaceBeforePoints", 0], ["paragraphSpaceAfterPoints", 0], ["paragraphMarginLeftEmu", 0], ["paragraphIndentEmu", 0]]) {
    const leaf = paragraphEdit.pages[0].elements[0].nativeRef.leaves.find(l => l.kind === kind);
    assert.ok(leaf, kind); leaf.value = value;
  }
  const paragraphCandidate = await compilePpjWorkspace({ ...paragraphInput, program: Buffer.from(JSON.stringify(paragraphEdit)) }, { includePreviewScene: true });
  const paragraphPixelsAfter = await paragraphPixels("paragraph-edited", paragraphCandidate);
  assert.deepEqual(paragraphPixelsAfter.starts.map((y, i) => y - paragraphPixelsBefore.starts[i]), [-10, 5, -10]);
  assert.deepEqual(paragraphPixelsAfter.lefts.map((x, i) => x - paragraphPixelsBefore.lefts[i]), [-20, -30, 0]);
  assert.deepEqual(paragraphPixelsAfter.xs, [107.2, 107.2, 107.2]);
  assert.ok(Math.abs(paragraphPixelsAfter.ys[1] - paragraphPixelsAfter.ys[0] - 55) < 1e-9);
  const paragraphFresh = await projectPptxToPpj(paragraphCandidate.file, { sourceUri: "paragraph-edited.pptx", assetRootUri: "assets" });
  const paragraphStyle = JSON.parse(new TextDecoder().decode(paragraphFresh.programJson)).pages[0].elements[0].text.paragraphs[0].style;
  assert.equal(paragraphStyle.lineSpacing, 55);
  assert.equal(paragraphStyle.spaceBefore, 0);
  assert.equal(paragraphStyle.spaceAfter, 0);
  assert.equal(paragraphStyle.indent, 0);
  assert.equal(paragraphStyle.hanging, 0);
  const paragraphOldZip = await JSZip.loadAsync(paragraphSource), paragraphNewZip = await JSZip.loadAsync(paragraphCandidate.file);
  assert.deepEqual(Object.keys(paragraphOldZip.files).sort(), Object.keys(paragraphNewZip.files).sort());
  const paragraphChangedParts = [];
  for (const name of Object.keys(paragraphOldZip.files)) if (!paragraphOldZip.files[name].dir &&
    !Buffer.from(await paragraphOldZip.file(name).async("uint8array")).equals(Buffer.from(await paragraphNewZip.file(name).async("uint8array")))) paragraphChangedParts.push(name);
  assert.deepEqual(paragraphChangedParts, ["ppt/slides/slide1.xml"]);
  assert.deepEqual(paragraphSource, paragraphBefore);
  const outlineProgram = structuredClone(pairBase);
  outlineProgram.pages[0].elements = [{ id: "outline", type: "shape", geometry: { kind: "preset", preset: "rect" },
    frame: { x: 100, y: 100, width: 200, height: 100 },
    style: { fill: { type: "solid", color: "#FFFFFF" }, stroke: { color: "#FF0000", width: 4, dash: "dash", cap: "flat", join: "miter" } } }];
  async function outlinePixels(name, result, edited = false) {
    const painted = await savePaint(name, result);
    const shape = createPpjSceneView(result).pages[0].nodes.find(n => n.kind === "shape").native;
    assert.equal(shape.lineStyle, edited ? "dotted" : "dashed");
    assert.equal(shape.lineCap, edited ? "round" : "flat");
    assert.equal(shape.lineJoin, edited ? "bevel" : "miter");
    assert.ok(painted.pages[0].svg.includes(`stroke-linecap="${edited ? "round" : "butt"}" stroke-linejoin="${edited ? "bevel" : "miter"}" stroke-dasharray="${edited ? "4 12" : "16 12"}"`));
    const { data, info } = await sharp(Buffer.from(painted.pages[0].svg)).flatten({ background: "#FFFFFF" }).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    for (const [x, expected] of edited ? [[110, [255,255,255]], [118, [255,0,0]]] : [[110, [255,0,0]], [122, [255,255,255]]]) {
      const offset = (100 * info.width + x) * info.channels;
      assert.deepEqual([...data.subarray(offset, offset + 3)], expected);
    }
  }
  const outlineAuthored = await compileFormat(outlineProgram);
  await outlinePixels("outline-authored", outlineAuthored);
  const outlineSource = await withoutAuthoredSnapshot(outlineAuthored.file), outlineBefore = outlineSource.slice();
  const outlineProjection = await projectPptxToPpj(outlineSource, { sourceUri: "outline.pptx", assetRootUri: "assets" });
  const outlineInput = { program: outlineProjection.programJson, source: outlineSource, assets: outlineProjection.assets };
  const outlineNoop = await compilePpjWorkspace(outlineInput, { includePreviewScene: true });
  assert.deepEqual(outlineNoop.file, outlineSource);
  await outlinePixels("outline-noop", outlineNoop);
  const outlineEdit = JSON.parse(new TextDecoder().decode(outlineProjection.programJson));
  for (const [kind, value] of [["lineStyle", "dotted"], ["lineCap", "round"], ["lineJoin", "bevel"]]) {
    const leaf = outlineEdit.pages[0].elements[0].nativeRef.leaves.find(l => l.kind === kind);
    assert.ok(leaf, kind); leaf.value = value;
  }
  const outlineCandidate = await compilePpjWorkspace({ ...outlineInput, program: Buffer.from(JSON.stringify(outlineEdit)) }, { includePreviewScene: true });
  await outlinePixels("outline-edited", outlineCandidate, true);
  const outlineFresh = await projectPptxToPpj(outlineCandidate.file, { sourceUri: "outline-edited.pptx", assetRootUri: "assets" });
  const outlineLeaves = JSON.parse(new TextDecoder().decode(outlineFresh.programJson)).pages[0].elements[0].nativeRef.leaves;
  for (const [kind, value] of [["lineStyle", "dotted"], ["lineCap", "round"], ["lineJoin", "bevel"]]) assert.equal(outlineLeaves.find(l => l.kind === kind).value, value);
  const outlineOldZip = await JSZip.loadAsync(outlineSource), outlineNewZip = await JSZip.loadAsync(outlineCandidate.file);
  assert.deepEqual(Object.keys(outlineOldZip.files).sort(), Object.keys(outlineNewZip.files).sort());
  const outlineChangedParts = [];
  for (const name of Object.keys(outlineOldZip.files)) if (!outlineOldZip.files[name].dir &&
    !Buffer.from(await outlineOldZip.file(name).async("uint8array")).equals(Buffer.from(await outlineNewZip.file(name).async("uint8array")))) outlineChangedParts.push(name);
  assert.deepEqual(outlineChangedParts, ["ppt/slides/slide1.xml"]);
  assert.deepEqual(outlineSource, outlineBefore);
  const diagramProgram = structuredClone(pairBase);
  diagramProgram.design.styles.shape = [{ id: "cache-shape", style: { fill: { type: "solid", color: "#CC5500" } } }];
  diagramProgram.design.styles.text = [{ id: "cache-text", style: { defaultText: { fontFamily: "Arial", size: 20, color: "#000000" } } }];
  diagramProgram.pages[0].elements = [{ id: "process", type: "smartArt", mode: "authored", layout: "process",
    frame: { x: 100, y: 100, width: 600, height: 180 }, shapeStyleRef: "cache-shape", textStyleRef: "cache-text",
    nodeGeometry: { kind: "preset", preset: "rect" },
    connector: { stroke: { color: "#000000", width: 2 }, endArrow: "triangle" },
    nodes: [{ id: "a", text: "Observe" }, { id: "b", text: "Decide" }],
    connections: [{ id: "edge", from: "a", to: "b", role: "sequence", order: 0 }] }];
  diagramProgram.pages[0].elements.push({ id: "diagram-sibling", type: "shape", frame: { x: 750, y: 100, width: 80, height: 80 },
    geometry: { kind: "preset", preset: "ellipse" }, style: { fill: { type: "solid", color: "#0088CC" } }, text: "Keep sibling" });
  async function diagramPixels(name, result, label) {
    const view = createPpjSceneView(result), diagram = view.pages[0].nodes.find(n => n.kind === "diagram");
    assert.ok(diagram?.native.drawingCacheVerified);
    if (result.previewScene.origin === 2) {
      const painted = await savePaint(name, result);
      assert.ok(diagram.native.nodes.some(n => n.textBody.paragraphs.some(p => p.runs.some(r => r.content.value === label))));
      assert.equal(painted.reliability.status, "failed");
      assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.diagram-import-incomplete"));
      assert.match(painted.pages[0].svg, /diagram: imported drawing incomplete/);
      return;
    }
    assert.ok(diagram.drawing.children.some(n => n.kind === "connector"));
    const shape = diagram.drawing.children.find(n => n.kind === "shape" && n.native.text.includes(label));
    assert.ok(shape, "Verified cache must contain current candidate text");
    const painted = await savePaint(name, result);
    assert.match(painted.pages[0].svg, /data-officekit-diagram="verified-cache"/);
    assert.ok(painted.pages[0].svg.includes(label));
    assert.ok(painted.pages[0].svg.includes(`data-officekit-native-id="${shape.nativeId}"`));
    const f = diagram.drawing.frame, c = diagram.drawing.childFrame;
    const x = Math.floor(f.x + (shape.frame.x + shape.frame.width / 2 - c.x) * f.width / c.width);
    const y = Math.floor(f.y + (shape.frame.y + shape.frame.height - 8 - c.y) * f.height / c.height);
    const { data, info } = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const offset = (y * info.width + x) * info.channels;
    assert.deepEqual([...data.subarray(offset, offset + 3)], [204, 85, 0]);
  }
  const diagramAuthored = await compileFormat(diagramProgram);
  await diagramPixels("diagram-authored", diagramAuthored, "Observe");
  const diagramSource = await withoutAuthoredSnapshot(diagramAuthored.file), diagramBefore = diagramSource.slice();
  const diagramProjection = await projectPptxToPpj(diagramSource, { sourceUri: "diagram.pptx", assetRootUri: "assets" });
  const diagramInput = { program: diagramProjection.programJson, source: diagramSource, assets: diagramProjection.assets };
  const diagramNoop = await compilePpjWorkspace(diagramInput, { includePreviewScene: true });
  assert.deepEqual(diagramNoop.file, diagramSource);
  await diagramPixels("diagram-noop", diagramNoop, "Observe");
  const diagramEdit = JSON.parse(new TextDecoder().decode(diagramProjection.programJson));
  diagramEdit.pages[0].elements.find(e => e.type === "smartArt").nodes[0].text = "Updated observation";
  const diagramCandidate = await compilePpjWorkspace({ ...diagramInput, program: Buffer.from(JSON.stringify(diagramEdit)) }, { includePreviewScene: true });
  await diagramPixels("diagram-edited", diagramCandidate, "Updated observation");
  const diagramFresh = await projectPptxToPpj(diagramCandidate.file, { sourceUri: "diagram-edited.pptx", assetRootUri: "assets" });
  assert.ok(new TextDecoder().decode(diagramFresh.programJson).includes("Updated observation"));
  const diagramOldZip = await JSZip.loadAsync(diagramSource), diagramNewZip = await JSZip.loadAsync(diagramCandidate.file);
  const diagramFailures = [];
  // This controlled fixture has exactly one SmartArt. Its source editor
  // replaces an exclusively owned graph, not a single drawing XML leaf.
  // Resolve actual relationship targets instead of allowing a directory glob.
  async function diagramGraph(zip) {
    const slide = await zip.file("ppt/slides/slide1.xml").async("string");
    const frames = [...slide.matchAll(/<(?:\w+:)?graphicFrame\b[\s\S]*?<\/(?:\w+:)?graphicFrame>/g)].map(m => m[0]);
    assert.equal(frames.length, 1);
    const ids = xmlAttributes(frames[0].match(/<(?:\w+:)?relIds\b[^>]*>/)[0]);
    const rels = [...(await zip.file("ppt/slides/_rels/slide1.xml.rels").async("string")).matchAll(/<Relationship\b[^>]*\/>/g)].map(m => xmlAttributes(m[0]));
    const targets = {}, relationshipIds = new Set();
    const resolve = (id, role) => {
      const rel = rels.find(r => r.Id === id);
      assert.ok(rel && !rel.TargetMode && rel.Type.endsWith(`/${role}`), `${role} relationship`);
      const part = path.posix.normalize(rel.Target.startsWith("/") ? rel.Target.slice(1) : path.posix.join("ppt/slides", rel.Target));
      assert.ok(zip.file(part), part); relationshipIds.add(id); return part;
    };
    for (const [key, role] of [["dm", "diagramData"], ["lo", "diagramLayout"], ["qs", "diagramQuickStyle"], ["cs", "diagramColors"]])
      targets[role] = resolve(Object.entries(ids).find(([name]) => name.endsWith(`:${key}`))[1], role);
    const data = await zip.file(targets.diagramData).async("string");
    const drawingId = xmlAttributes(data.match(/<(?:\w+:)?dataModelExt\b[^>]*>/)[0]).relId;
    targets.diagramDrawing = resolve(drawingId, "diagramDrawing");
    return { slide, frame: frames[0], rels, relationshipIds, targets };
  }
  const diagramOldGraph = await diagramGraph(diagramOldZip), diagramNewGraph = await diagramGraph(diagramNewZip);
  const diagramOwnedParts = new Set([...Object.values(diagramOldGraph.targets), ...Object.values(diagramNewGraph.targets)]);
  const diagramOwnedRelationships = new Set([...diagramOldGraph.relationshipIds, ...diagramNewGraph.relationshipIds]);
  const diagramChangedParts = [];
  for (const name of new Set([...Object.keys(diagramOldZip.files), ...Object.keys(diagramNewZip.files)])) {
    const old = diagramOldZip.file(name), candidate = diagramNewZip.file(name);
    if (!old && !candidate) continue;
    if (!old || !candidate || !Buffer.from(await old.async("uint8array")).equals(Buffer.from(await candidate.async("uint8array")))) diagramChangedParts.push(name);
  }
  assert.ok(diagramChangedParts.length > 0);
  try {
    const permitted = new Set([...diagramOwnedParts, "[Content_Types].xml", "ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels"]);
    assert.ok(diagramChangedParts.every(name => permitted.has(name)), JSON.stringify(diagramChangedParts));
    // Whole other files are byte-equal, because every changed/added/removed
    // part was classified above. Check shared parts below at subtree scope.
    const outsideBefore = diagramOldGraph.slide.replace(diagramOldGraph.frame, "TARGET"), outsideAfter = diagramNewGraph.slide.replace(diagramNewGraph.frame, "TARGET");
    assert.equal(orderedXml(outsideBefore), orderedXml(outsideAfter));
    assert.notEqual(orderedXml(outsideBefore), orderedXml(outsideAfter.replace("Keep sibling", "WRONG sibling")));
    const unrelatedRels = graph => graph.rels.filter(r => !diagramOwnedRelationships.has(r.Id)).sort((a,b) => a.Id.localeCompare(b.Id));
    assert.deepEqual(unrelatedRels(diagramOldGraph), unrelatedRels(diagramNewGraph));
    const unrelatedTypes = async zip => [...(await zip.file("[Content_Types].xml").async("string")).matchAll(/<(?:Default|Override)\b[^>]*\/>/g)]
      .map(m => xmlAttributes(m[0])).filter(a => !diagramOwnedParts.has(a.PartName?.replace(/^\//, "")))
      .sort((a,b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
    assert.deepEqual(await unrelatedTypes(diagramOldZip), await unrelatedTypes(diagramNewZip));
    for (const role of ["diagramLayout", "diagramQuickStyle", "diagramColors"])
      assert.deepEqual(await diagramOldZip.file(diagramOldGraph.targets[role]).async("uint8array"), await diagramNewZip.file(diagramNewGraph.targets[role]).async("uint8array"));
  } catch (error) { diagramFailures.push(error); }
  await writeFile(path.join(artifacts, "diagram-source.pptx"), diagramSource, { flag: "wx" });
  await writeFile(path.join(artifacts, "diagram-candidate.pptx"), diagramCandidate.file, { flag: "wx" });
  assert.deepEqual(diagramSource, diagramBefore);
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
  const nestedFixture = JSON.parse(await readFile(new URL("./fixtures/presentation/preview-nested-repeat-equivalence.json", import.meta.url)));
  assert.equal(nestedFixture.base, "examples/ppj/minimum.ppj");
  const styleFixture = JSON.parse(await readFile(new URL("./fixtures/presentation/preview-style-grammar-equivalence.json", import.meta.url)));
  const nestedPairFailures = [];
  const nativeContentSchemas = new Map(PresentationElementSchema.fields.filter(f => f.oneof?.localName === "content").map(f => [f.localName, f.message]));
  for (const variant of ["plain", "styled"]) {
  const nestedHighLevel = structuredClone(pairBase), nestedExplicit = structuredClone(pairBase);
  nestedHighLevel.components = structuredClone(nestedFixture.components);
  nestedHighLevel.pages[0].elements = [nestedFixture.instance, nestedFixture.risk];
  nestedExplicit.pages[0].elements = structuredClone([...nestedFixture.explicit, nestedFixture.risk]);
  if (variant === "styled") {
    // Compose two independently specified fixtures, never derive explicit
    // geometry or style from compiler output.
    nestedHighLevel.design.styles = styleFixture.styles;
    nestedHighLevel.design.grammar.tokens = styleFixture.tokens;
    const tile = nestedHighLevel.components[0].elements[0];
    delete tile.style;
    tile.styleRef = styleFixture.named[0].styleRef;
    const label = nestedHighLevel.components[1].elements[0].slots.label[0];
    label.styleRef = styleFixture.named[1].styleRef;
    label.style = styleFixture.named[1].style;
    nestedExplicit.design.fonts.push({ id: "resolved-face", family: "DejaVu Sans" });
    for (let i = 0; i < 4; i++) nestedExplicit.pages[0].elements[i].style = structuredClone(styleFixture.explicit[i % 2].style);
  }
  const nestedPair = [];
  for (const [name, program] of [["component", nestedHighLevel], ["explicit", nestedExplicit]]) {
    try {
    const original = JSON.stringify(program);
    await writeFile(path.join(artifacts, `nested-pair-${variant}-${name}.ppj`), original, { flag: "wx" });
    const ordinary = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(original) });
    assert.equal(Object.hasOwn(ordinary, "previewScene"), false);
    const compiled = await compileFormat(program), painted = await savePaint(`nested-pair-${variant}-${name}`, compiled);
    assert.deepEqual(compiled.file, ordinary.file, "collecting repeat/slot origins must not change the exported candidate");
    const scene = createPpjSceneView(compiled), nodes = scene.pages[0].nodes;
    assert.equal(nodes.length, 5);
    for (let i = 0; i < 4; i++) {
      assert.equal(nodes[i].kind, "shape");
      assert.deepEqual(nodes[i].frame, nestedFixture.explicit[i].frame);
      if (i % 2 === 0) assert.equal(nodes[i].native.fillRgb, "CC5500");
      else assert.equal(nodes[i].native.text, "0");
      if (variant === "styled" && i % 2 === 1) {
        const runs = nodes[i].native.textBody.paragraphs.flatMap(p => p.runs);
        assert.equal(runs.length, 1);
        assert.equal(runs[0].fontFamily, "DejaVu Sans");
        assert.equal(runs[0].fontSizePoints, 32);
        assert.equal(runs[0].bold, false);
      }
    }
    assert.equal(nodes[4].kind, "chart");
    assert.deepEqual(nodes[4].native.series[0].values, [1, 0, 0]);
    assert.deepEqual(nodes[4].native.series[0].missingValueIndexes, [1]);
    assert.equal((painted.pages[0].svg.match(/data-officekit-review-point="isolated"/g) || []).length, 2);
    assert.doesNotMatch(painted.pages[0].svg, /data-officekit-line-segment=/);
    if (name === "component") {
      const tiles = compiled.previewScene.bindings.filter(b => b.sourceId === "tile");
      const labels = compiled.previewScene.bindings.filter(b => b.sourceId === "supplied-zero");
      assert.equal(tiles.length, 2); assert.equal(labels.length, 2);
      for (const binding of tiles) assert.equal(binding.programPath, nestedFixture.expected.tileOwner);
      for (const binding of labels) assert.equal(binding.programPath, nestedFixture.expected.slotOwner);
      assert.equal(new Set(tiles.map(b => b.instanceId)).size, 2);
      assert.equal(new Set([...tiles, ...labels].map(b => b.scenePath)).size, 4);
      assert.ok([...tiles, ...labels].every(b => b.attribution === 2));
    }
    const raster = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * raster.info.channels, (y * raster.info.width + x) * raster.info.channels + 3)];
    for (const point of nestedFixture.expected.orangePixels) assert.deepEqual(pixel(...point), [204, 85, 0]);
    for (const point of nestedFixture.expected.blankPixels) assert.deepEqual(pixel(...point), [255, 255, 255]);
    if (variant === "styled") for (const left of [150, 350]) {
      let ink = 0;
      // Font size exceeds the small slot height: inspect actual glyph extent,
      // without treating this equivalence case as an overflow-layout pass.
      for (let y = 110; y < 155; y++) for (let x = left; x < left + 40; x++) {
        if (pixel(x, y).every((v, i) => v === [17, 68, 119][i])) ink++;
      }
      assert.ok(ink > 20, "each repeated slot must paint its resolved grammar color");
    }
    assert.equal(JSON.stringify(program), original);
    // Compare every typed visual payload byte. Wrapper IDs/provenance differ
    // by construction and are asserted separately, not normalized into PPJ.
    nestedPair.push({ visual: nodes.map(node => ({ kind: node.kind, hidden: node.hidden,
      bytes: toBinary(nativeContentSchemas.get(node.kind), node.native) })), raster: raster.data });
    await assertProductionEntry(`nested-${variant}-${name}`, { ...sourceWorkspace, program: Buffer.from(original) }, compiled, painted);
    } catch (error) {
      nestedPairFailures.push({ variant, name, code: error.code, message: error.message });
      console.error(`Nested repeat pair ${name} failed: ${error.message}`);
    }
  }
  if (nestedPair.length === 2) try {
    assert.deepEqual(nestedPair[0].visual, nestedPair[1].visual);
    assert.deepEqual(nestedPair[0].raster, nestedPair[1].raster);
  } catch (error) { nestedPairFailures.push({ variant, name: "equivalence", code: error.code, message: error.message }); }
  }
  const datasetFixture = JSON.parse(await readFile(new URL("./fixtures/presentation/preview-dataset-equivalence.json", import.meta.url)));
  const datasetPairFailures = [];
  const datasetProfileCases = [];
  for (const [chartType, firstValue] of [["line", 1], ["line", 2], ["heatmap", 1]]) {
  const variant = firstValue === 1 ? chartType : `${chartType}-value-${firstValue}`;
  const datasetPair = [];
  for (const name of ["encoded", "explicit"]) try {
    const program = structuredClone(pairBase);
    program.pages[0].elements = [{ id: "dataset-line", type: "chart", chartType, frame: datasetFixture.frame, data: datasetFixture[name] }];
    program.pages[0].elements[0].data = structuredClone(datasetFixture[name]);
    if (name === "encoded") program.pages[0].elements[0].data.dataset.rows[0][1] = firstValue;
    else program.pages[0].elements[0].data.series[0].values[0] = firstValue;
    if (chartType === "heatmap") program.pages[0].elements[0].style = { heatmap: {
      colors: ["#000000", "#FF0000"], domain: [0, 5], missingFill: "#00FF00", showColorBar: false, showValues: false, cellGap: 0,
    } };
    const original = JSON.stringify(program);
    await writeFile(path.join(artifacts, `dataset-pair-${variant}-${name}.ppj`), original, { flag: "wx" });
    const ordinary = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(original) });
    const compiled = await compileFormat(program), painted = await savePaint(`dataset-pair-${variant}-${name}`, compiled);
    assert.deepEqual(compiled.file, ordinary.file);
    const node = createPpjSceneView(compiled).pages[0].nodes[0];
    assert.deepEqual(node.frame, datasetFixture.frame);
    if (chartType === "line") {
    assert.equal(node.kind, "chart");
    assert.deepEqual(node.native.series.map(s => s.name), ["Alpha", "Beta"]);
    assert.deepEqual(node.native.series.map(s => s.values), [[firstValue, 0, 0], [0, 4, 5]]);
    assert.deepEqual(node.native.series.map(s => s.missingValueIndexes), [[1], []]);
    assert.equal((painted.pages[0].svg.match(/data-officekit-review-point="isolated"/g) || []).length, 2);
    assert.equal((painted.pages[0].svg.match(/data-officekit-line-segment=/g) || []).length, 1);
    const pixels = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const blueNear = (x, y) => {
      let count = 0;
      for (let py = y - 5; py <= y + 5; py++) for (let px = x - 5; px <= x + 5; px++) {
        const offset = (py * pixels.info.width + px) * pixels.info.channels;
        if (pixels.data[offset] < 100 && pixels.data[offset + 1] < 160 && pixels.data[offset + 2] > 150) count++;
      }
      return count;
    };
    assert.ok(blueNear(160, 355 - firstValue * 42) > 0, "changed source value moves the actual isolated-point pixels");
    assert.equal(blueNear(160, 355 - (firstValue === 1 ? 2 : 1) * 42), 0);
    assert.ok(blueNear(640, 355) > 0, "real zero remains visible");
    assert.equal(blueNear(400, 355), 0, "missing category must not become a zero point");
    } else {
      assert.equal(node.kind, "group", "vector heatmap must consume the compiler's actual children");
      const cells = node.native.children.filter(c => c.name.startsWith("heatmap cell "));
      assert.equal(cells.length, 6);
      const fills = ["330000", "00FF00", "000000", "000000", "CC0000", "FF0000"];
      assert.deepEqual(cells.map(c => c.content.value.fillRgb), fills);
      const raster = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
      // Fixed expected layout: 52pt row-label gutter, 27pt bottom labels,
      // three columns and two rows in the 600x300 frame.
      for (const [i, cell] of cells.entries()) {
        const shape = cell.content.value, row = Math.floor(i / 3), col = i % 3;
        assert.ok(Math.abs(Number(shape.leftEmu) / 12700 - (152 + col * 548 / 3)) < 0.0001);
        assert.equal(Number(shape.topEmu) / 12700, 100 + row * 136.5);
        assert.ok(Math.abs(Number(shape.widthEmu) / 12700 - 548 / 3) < 0.0001);
        assert.equal(Number(shape.heightEmu) / 12700, 136.5);
        const x = Math.floor(152 + (col + 0.5) * 548 / 3), y = Math.floor(100 + (row + 0.5) * 136.5);
        const offset = (y * raster.info.width + x) * raster.info.channels;
        assert.deepEqual([...raster.data.subarray(offset, offset + 3)], fills[i].match(/../g).map(v => parseInt(v, 16)));
      }
      assert.equal(new Set(compiled.previewScene.bindings.map(b => b.scenePath)).size, compiled.previewScene.bindings.length);
      assert.ok(compiled.previewScene.bindings.length > 6, "generated children retain individual scene addresses");
    }
    assert.ok(compiled.previewScene.bindings.every(b => b.programPath === "$.pages[0].elements[0]" && b.semanticId === "dataset-line"));
    const raster = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer();
    assert.equal(JSON.stringify(program), original);
    datasetPair.push({ visual: toBinary(nativeContentSchemas.get(node.kind), node.native), raster });
    const published = await assertProductionEntry(`dataset-${variant}-${name}`, { ...sourceWorkspace, program: Buffer.from(original) }, compiled, painted);
    if (name === "encoded") {
      const canonical = JSON.parse(new TextDecoder().decode(compiled.programJson));
      const field = "$.pages[0].elements[0].data.dataset", reason = "preview.fact.chart-channel-ignored";
      const isDatasetError = d => d.path === field && d.reason === reason;
      const legacy = assessPpjPreviewInput(canonical);
      const options = { rendererProfile: "native-scene-svg", sceneReceipt: compiled, scenePaint: painted };
      const mapped = assessPpjPreviewInput(canonical, options);
      assert.ok(legacy.diagnostics.some(isDatasetError));
      assert.equal(mapped.diagnostics.some(isDatasetError), chartType !== "line");
      assert.equal(published.diagnostics.some(isDatasetError), chartType !== "line");
      for (const d of legacy.diagnostics.filter(d => !isDatasetError(d)))
        assert.ok(mapped.diagnostics.some(a => a.path === d.path && a.reason === d.reason), "unrelated input checks remain");
      const missing = assessPpjPreviewInput(canonical, { ...options, scenePaint: { ...painted, lineScenePaths: [] } });
      assert.ok(missing.diagnostics.some(isDatasetError));
      const registry = JSON.parse(await readFile("src/ppj/capability-registry.json", "utf8"));
      delete registry.previewScene.factualMappings.datasetLine;
      assert.throws(() => assessPpjPreviewInput(canonical, { ...options, registry }), /dataset line mapping/);
      datasetProfileCases.push({ chartType, firstValue, errorRetired: chartType === "line", missingCaptureRetainsError: true,
        missingRegistryRejects: true, unrelatedRulesRetained: true, productionReliability: published.reliability.status,
        candidateSha256: sha256(compiled.file), sceneSha256: compiled.previewScene.sha256 });
    }
  } catch (error) {
    datasetPairFailures.push({ chartType, firstValue, name, code: error.code, message: error.message });
    console.error(`Dataset pair ${name} failed: ${error.message}`);
  }
  if (datasetPair.length === 2) try {
    assert.deepEqual(datasetPair[0].visual, datasetPair[1].visual);
    assert.deepEqual(datasetPair[0].raster, datasetPair[1].raster);
  } catch (error) { datasetPairFailures.push({ chartType, firstValue, name: "equivalence", code: error.code, message: error.message }); }
  }
  assert.equal(styleFixture.base, "examples/ppj/minimum.ppj");
  const stylePair = [], stylePairFailures = [];
  for (const name of ["named", "explicit"]) try {
    const program = structuredClone(pairBase);
    program.design.fonts.push({ id: "resolved-face", family: "DejaVu Sans" });
    if (name === "named") {
      program.design.styles = styleFixture.styles;
      program.design.grammar.tokens = styleFixture.tokens;
    }
    program.pages[0].elements = [...styleFixture[name], styleFixture.risk];
    const original = JSON.stringify(program);
    await writeFile(path.join(artifacts, `style-pair-${name}.ppj`), original, { flag: "wx" });
    const ordinary = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(original) });
    const compiled = await compileFormat(program), painted = await savePaint(`style-pair-${name}`, compiled);
    assert.equal(Object.hasOwn(ordinary, "previewScene"), false);
    assert.deepEqual(compiled.file, ordinary.file);
    const nodes = createPpjSceneView(compiled).pages[0].nodes;
    assert.deepEqual(nodes.map(n => n.kind), ["shape", "shape", "chart"]);
    assert.equal(nodes[0].native.fillRgb, "CC5500");
    const runs = nodes[1].native.textBody.paragraphs.flatMap(p => p.runs);
    assert.equal(runs.length, 1);
    assert.equal(runs[0].fontFamily, "DejaVu Sans");
    assert.equal(runs[0].fontSizePoints, 32);
    assert.equal(runs[0].bold, false, "explicit false overrides named true");
    assert.deepEqual(nodes[2].native.series[0].values, [1, 0, 0]);
    assert.deepEqual(nodes[2].native.series[0].missingValueIndexes, [1]);
    assert.equal((painted.pages[0].svg.match(/data-officekit-review-point="isolated"/g) || []).length, 2);
    assert.doesNotMatch(painted.pages[0].svg, /data-officekit-line-segment=/);
    assert.match(painted.pages[0].svg, /font-family="DejaVu Sans"/);
    assert.match(painted.pages[0].svg, /font-size="32"/);
    for (const [i, node] of nodes.entries()) {
      assert.deepEqual(node.frame, program.pages[0].elements[i].frame);
      assert.ok(compiled.previewScene.bindings.some(b => b.semanticId === program.pages[0].elements[i].id && b.programPath === `$.pages[0].elements[${i}]`));
    }
    const raster = await sharp(Buffer.from(painted.pages[0].svg)).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * raster.info.channels, (y * raster.info.width + x) * raster.info.channels + 3)];
    assert.deepEqual(pixel(150, 130), [204, 85, 0]);
    let ink = 0;
    for (let y = 100; y < 170; y++) for (let x = 230; x < 530; x++) {
      if (pixel(x, y).every((v, i) => v === [17, 68, 119][i])) ink++;
    }
    assert.ok(ink > 100, "resolved text color must reach actual glyph pixels");
    assert.equal(JSON.stringify(program), original);
    stylePair.push({ visual: nodes.map(n => ({ kind: n.kind, bytes: toBinary(nativeContentSchemas.get(n.kind), n.native) })), raster: raster.data });
  } catch (error) {
    stylePairFailures.push({ name, code: error.code, message: error.message });
    console.error(`Style/grammar pair ${name} failed: ${error.message}`);
  }
  if (stylePair.length === 2) try {
    assert.deepEqual(stylePair[0].visual, stylePair[1].visual);
    assert.deepEqual(stylePair[0].raster, stylePair[1].raster);
  } catch (error) { stylePairFailures.push({ name: "equivalence", code: error.code, message: error.message }); }
  const scatterProgram = structuredClone(pairBase);
  scatterProgram.pages[0].elements = [{ id: "numeric", type: "chart", chartType: "scatter",
    frame: { x: 100, y: 100, width: 500, height: 300 }, style: { scatterStyle: "marker", legend: "none" },
    xAxis: { min: 0, max: 100 }, yAxis: { min: 0, max: 10 },
    data: { categories: [], series: [{ id: "xy", name: "XY", xValues: [0, 10, 50, 100], values: [0, 2, null, 10],
      marker: { symbol: "circle", size: 8, fill: "#FF0000" } }] } }];
  async function scatterPixels(name, receipt, middleX = 10, middleY = 2) {
    const native = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "chart").native;
    assert.equal(native.type, 6); assert.equal(native.scatterStyle, "marker");
    assert.deepEqual(native.series[0].xValues, [0, middleX, 50, 100]);
    assert.deepEqual(native.series[0].values, [0, middleY, 0, 10]);
    assert.deepEqual(native.series[0].missingValueIndexes, [2]);
    const painted = await savePaint(name, receipt), svg = painted.pages[0].svg;
    assert.equal(painted.reliability.status, "requires-review", JSON.stringify(painted.diagnostics));
    assert.match(svg, new RegExp(`circle cx="${150 + middleX * 4}" cy="${355 - middleY * 21}" r="4"`));
    assert.doesNotMatch(svg, /data-officekit-line-segment=/);
    assert.doesNotMatch(svg, /data-officekit-point="2"|data-officekit-line-segment="0:3"/);
    const raster = await sharp(Buffer.from(svg)).flatten({ background: "#FFFFFF" }).removeAlpha().raw().toBuffer({ resolveWithObject: true });
    const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * raster.info.channels, (y * raster.info.width + x) * raster.info.channels + 3)];
    assert.deepEqual(pixel(150 + middleX * 4, 355 - middleY * 21), [255, 0, 0]);
    assert.deepEqual(pixel(350, 353), [255, 255, 255], "Missing Y must not create a zero marker above the axis");
    assert.deepEqual(pixel(550, 145), [255, 0, 0], "Isolated last observation remains a marker");
  }
  const scatterAuthored = await compileFormat(scatterProgram);
  await scatterPixels("scatter-authored", scatterAuthored);
  const scatterAuthoredX = structuredClone(scatterProgram);
  scatterAuthoredX.pages[0].elements[0].data.series[0].xValues[1] = 25;
  await scatterPixels("scatter-authored-x", await compileFormat(scatterAuthoredX), 25);
  const scatterConnectedProgram = structuredClone(scatterProgram);
  scatterConnectedProgram.pages[0].elements[0].style.scatterStyle = "lineWithMarkers";
  const scatterConnected = await compileFormat(scatterConnectedProgram);
  const scatterConnectedPaint = await savePaint("scatter-connected-unavailable", scatterConnected);
  assert.equal(scatterConnectedPaint.reliability.status, "failed");
  assert.ok(scatterConnectedPaint.diagnostics.some(d => d.reason === "preview.scene.paint.scatter-line-unresolved"));
  assert.doesNotMatch(scatterConnectedPaint.pages[0].svg, /data-officekit-line-segment=/);
  const scatterConnectedZip = await JSZip.loadAsync(scatterConnected.file);
  const scatterConnectedXml = await scatterConnectedZip.file("ppt/slides/charts/chart1.xml").async("string");
  assert.match(scatterConnectedXml, /scatterStyle[^>]*val="lineMarker"/);
  assert.match(scatterConnectedXml, /<a:ln><a:noFill\s*\/><\/a:ln>/);
  const scatterSource = await withoutAuthoredSnapshot(scatterAuthored.file), scatterBefore = scatterSource.slice();
  const scatterProjection = await projectPptxToPpj(scatterSource, { sourceUri: "scatter.pptx", assetRootUri: "assets" });
  const scatterInput = { program: scatterProjection.programJson, source: scatterSource, assets: scatterProjection.assets };
  const scatterNoop = await compilePpjWorkspace(scatterInput, { includePreviewScene: true });
  assert.deepEqual(scatterNoop.file, scatterSource); await scatterPixels("scatter-noop", scatterNoop);
  const scatterEdit = JSON.parse(new TextDecoder().decode(scatterProjection.programJson));
  scatterEdit.pages[0].elements[0].data.series[0].xValues[1] = 25;
  await assert.rejects(compilePpjWorkspace({ ...scatterInput, program: Buffer.from(JSON.stringify(scatterEdit)) }, { includePreviewScene: true }),
    error => error.code === "ppj.source.unsupportedMutation" && /xValues/.test(error.message));
  scatterEdit.pages[0].elements[0].data.series[0].xValues[1] = 10;
  scatterEdit.pages[0].elements[0].data.series[0].values[1] = 5;
  const scatterCandidate = await compilePpjWorkspace({ ...scatterInput, program: Buffer.from(JSON.stringify(scatterEdit)) }, { includePreviewScene: true });
  await scatterPixels("scatter-edited", scatterCandidate, 10, 5);
  const scatterFresh = await projectPptxToPpj(scatterCandidate.file, { sourceUri: "scatter-edited.pptx", assetRootUri: "assets" });
  const scatterFreshSeries = JSON.parse(new TextDecoder().decode(scatterFresh.programJson)).pages[0].elements[0].data.series[0];
  assert.deepEqual(scatterFreshSeries.xValues, [0, 10, 50, 100]);
  assert.deepEqual(scatterFreshSeries.values, [0, 5, null, 10]);
  const scatterOldZip = await JSZip.loadAsync(scatterSource), scatterNewZip = await JSZip.loadAsync(scatterCandidate.file), scatterChangedParts = [];
  assert.deepEqual(Object.keys(scatterOldZip.files).sort(), Object.keys(scatterNewZip.files).sort());
  for (const name of Object.keys(scatterOldZip.files)) if (!scatterOldZip.files[name].dir &&
    !Buffer.from(await scatterOldZip.file(name).async("uint8array")).equals(Buffer.from(await scatterNewZip.file(name).async("uint8array")))) scatterChangedParts.push(name);
  assert.deepEqual(scatterChangedParts, ["ppt/slides/charts/chart1.xml"]);
  assert.deepEqual(scatterSource, scatterBefore);
  const isolatedLineProfileCases = [];
  const isolatedLineRegistry = JSON.parse(await readFile("src/ppj/capability-registry.json", "utf8"));
  function assertIsolatedLineProfile(receipt, painted) {
    const reasons = new Set(["preview.fact.missing-observation-misrepresented", "preview.fact.series-type-not-inherited"]);
    const program = JSON.parse(Buffer.from(receipt.programJson).toString("utf8"));
    const before = sha256(receipt.file), canonical = assessPpjPreviewInput(program);
    const errors = canonical.diagnostics.filter(d => reasons.has(d.reason));
    assert.ok(errors.length > 0, "Original line-profile factual error must be exercised");
    const options = { rendererProfile: "native-scene-svg", sceneReceipt: receipt, scenePaint: painted };
    const actual = assessPpjPreviewInput(program, options);
    assert.ok(!actual.diagnostics.some(d => reasons.has(d.reason)));
    const missing = assessPpjPreviewInput(program, { ...options, scenePaint: { ...painted, isolatedLinePoints: [] } });
    assert.deepEqual(missing.diagnostics.filter(d => reasons.has(d.reason)), errors);
    // Bundled source fixtures also exercise the independently verified direct
    // connector mapping below; it is no longer an unchanged legacy rule.
    for (const diagnostic of canonical.diagnostics.filter(d => !reasons.has(d.reason) && d.reason !== "preview.fact.connector-endpoints-ignored"))
      assert.ok(actual.diagnostics.some(d => d.reason === diagnostic.reason && d.path === diagnostic.path && d.severity === diagnostic.severity));
    const registry = structuredClone(isolatedLineRegistry);
    delete registry.previewScene.factualMappings.isolatedLinePoints;
    assert.throws(() => assessPpjPreviewInput(program, { ...options, registry }), /isolated line point mapping/);
    const combined = paintPpjSceneSvg(receipt, { assessInput: true });
    assert.ok(!combined.diagnostics.some(d => reasons.has(d.reason)));
    for (const diagnostic of painted.diagnostics)
      assert.ok(combined.diagnostics.some(d => d.reason === diagnostic.reason && d.scenePath === diagnostic.scenePath));
    assert.equal(sha256(receipt.file), before);
    isolatedLineProfileCases.push({ origin: receipt.previewScene.origin, retiredErrors: errors.map(d => ({ reason: d.reason, path: d.path })),
      points: painted.isolatedLinePoints, candidateSha256: before });
  }
  const nativeLineInput = { id: "native-line", type: "chart", chartType: "line", title: "Missing is not zero",
    frame: { x: 500, y: 330, width: 400, height: 160 }, yAxis: { min: 0, max: 5 },
    data: { categories: ["A", "Missing", "Zero", "D", "E"], series: [{ id: "observed", name: "Observed",
      values: [2, null, 0, 4, 5], stroke: { color: "#047857", width: 2 },
      marker: { symbol: "circle", size: 8, fill: "#047857", stroke: { color: "#114477", width: 2 } } },
      { id: "other", name: "Other series", values: [5, 4, null, 0, 1], stroke: { color: "#B45309", width: 2 }, marker: "none" }] } };
  async function assertNativeLine(receipt, painted, middle = 0) {
    assertIsolatedLineProfile(receipt, painted);
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
  const connectorProfileCases = [];
  const connectorRegistry = JSON.parse(await readFile(new URL("../src/ppj/capability-registry.json", import.meta.url)));
  function assertDirectEndpointProfile(receipt, painted) {
    const reason = "preview.fact.connector-endpoints-ignored";
    const program = JSON.parse(Buffer.from(receipt.programJson).toString("utf8"));
    const before = sha256(receipt.file), canonical = assessPpjPreviewInput(program);
    const options = { rendererProfile: "native-scene-svg", sceneReceipt: receipt, scenePaint: painted };
    const actual = assessPpjPreviewInput(program, options);
    assert.equal(canonical.diagnostics.filter(d => d.reason === reason).length, 2);
    assert.ok(!actual.diagnostics.some(d => d.reason === reason));
    const missing = assessPpjPreviewInput(program, { ...options, scenePaint: { ...painted, connectorScenePaths: [] } });
    assert.equal(missing.diagnostics.filter(d => d.reason === reason).length, 2, "receipt alone is not actual route evidence");
    // These exact line rules are now checked separately by assertNativeLine,
    // including their missing-capture restoration and actual marker/gap pixels.
    for (const diagnostic of canonical.diagnostics.filter(d => d.reason !== reason &&
      !["preview.fact.missing-observation-misrepresented", "preview.fact.series-type-not-inherited"].includes(d.reason)))
      assert.ok(actual.diagnostics.some(d => d.reason === diagnostic.reason && d.path === diagnostic.path && d.severity === diagnostic.severity), "unrelated limitations must survive");
    const registry = structuredClone(connectorRegistry);
    delete registry.previewScene.factualMappings.connectorEndpoints;
    assert.throws(() => assessPpjPreviewInput(program, { ...options, registry }), /connector endpoint mapping/);
    const combined = paintPpjSceneSvg(receipt, { assessInput: true });
    assert.ok(!combined.diagnostics.some(d => d.reason === reason));
    for (const diagnostic of painted.diagnostics)
      assert.ok(combined.diagnostics.some(d => d.reason === diagnostic.reason && d.scenePath === diagnostic.scenePath));
    assert.equal(sha256(receipt.file), before);
    connectorProfileCases.push({ origin: receipt.previewScene.origin, sceneSha256: receipt.previewScene.sha256,
      candidateSha256: before, routeCaptures: painted.connectorScenePaths.length, retiredEndpoints: 2 });
  }
  function assertLiteralEdge(receipt, painted) {
    const edge = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "connector");
    assert.deepEqual(edge.endpoints, { start: { x: 750, y: 320 }, end: { x: 510, y: 120 } });
    assert.ok(painted.pages[0].svg.includes('data-officekit-connector="straight" d="M 750 320 L 510 120"'));
    assert.ok(painted.pages[0].svg.includes('data-officekit-arrow="end" data-officekit-arrow-kind="triangle" transform="translate(510 120) rotate(-'));
    assertDirectEndpointProfile(receipt, painted);
  }
  const literalProgram = structuredClone(pairBase);
  const cropProgram = structuredClone(pairBase);
  const cropData = await sharp({ create: { width: 40, height: 20, channels: 4, background: "#FF0000" } })
    .composite([{ input: await sharp({ create: { width: 20, height: 20, channels: 4, background: "#00FF00" } }).png().toBuffer(), left: 20, top: 0 }]).png().toBuffer();
  const cropAsset = { ...JSON.parse(Buffer.from(authoredAssets.program).toString("utf8")).assets[0], id: "crop-image", mimeType: "image/png", sha256: sha256(cropData) };
  // A transparent central stripe distinguishes asset alpha from whole-fill
  // opacity; the same asset is also a foreground picture, so background edits
  // must preserve its bytes, relationship and visible foreground use.
  const backgroundData = await sharp({ create: { width: 60, height: 20, channels: 4, background: "#00000000" } })
    .composite([{ input: await sharp({ create: { width: 20, height: 20, channels: 4, background: "#FF0000" } }).png().toBuffer(), left: 0, top: 0 },
      { input: await sharp({ create: { width: 20, height: 20, channels: 4, background: "#00FF00" } }).png().toBuffer(), left: 40, top: 0 }]).png().toBuffer();
  const backgroundAsset = { ...cropAsset, id: "background-image", sha256: sha256(backgroundData) };
  const transparent = [0, 0, 0, 0], red = alpha => [255, 0, 0, alpha], green = alpha => [0, 255, 0, alpha];
  const stretchSamples = alpha => [[.1, red(alpha)], [.5, transparent], [.9, green(alpha)]];
  const croppedSamples = alpha => [[.1, transparent], [.5, green(alpha)], [.9, green(alpha)]];
  const letterboxSamples = [[.1, transparent], [.3, red(128)], [.5, transparent], [.7, green(128)], [.9, transparent]];
  async function backgroundImagePixels(name, receipt, expected, samples) {
    const painted = await savePaint(`background-image-${name}`, receipt), view = createPpjSceneView(receipt);
    const paint = view.pages[0].native.background?.imagePaint;
    if (expected === null) {
      assert.equal(paint, undefined);
      assert.doesNotMatch(painted.pages[0].svg, /data-officekit-background="image"/);
    } else {
      assert.equal(paint.mode, 1);
      assert.equal(paint.opacityThousandthPercent, expected.opacity === undefined ? undefined : expected.opacity * 100000);
      if (!expected.crop) assert.equal(paint.crop, undefined);
      else for (const side of ["left", "top", "right", "bottom"])
        assert.equal(paint.crop[`${side}ThousandthPercent`], (expected.crop[side] ?? 0) * 100000);
      assert.deepEqual(Buffer.from(view.asset(paint.assetId).data), backgroundData);
      assert.match(painted.pages[0].svg, /data-officekit-background="image"/);
    }
    const raster = await sharp(Buffer.from(painted.pages[0].svg)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
    const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * 4, (y * raster.info.width + x) * 4 + 4)];
    for (const [fraction, rgba] of samples)
      assert.deepEqual(pixel(Math.floor(fraction * raster.info.width), Math.floor(raster.info.height / 2)), rgba, `${name}: native image/crop/alpha pixels at ${fraction}`);
    assert.deepEqual(pixel(120, 120), [204, 85, 0, 255], "foreground shape is above the background");
    assert.deepEqual(pixel(205, 110), red(255), "shared foreground image retains its own full alpha");
    assert.deepEqual(pixel(255, 110), green(255));
    assert.ok(!painted.diagnostics.some(d => d.scenePath?.includes(".background")));
    backgroundImageCases.push({ name, nativeAndRgbaPixels: true, transparentAssetRetained: true, foregroundPreserved: true,
      candidateSha256: sha256(receipt.file), sceneSha256: receipt.previewScene.sha256 });
    return painted;
  }
  for (const [name, fill, samples] of [
    ["stretch", {}, stretchSamples(255)],
    ["zero", { opacity: 0 }, [[.1, transparent], [.5, transparent], [.9, transparent]]],
    ["crop", { opacity: .5, crop: { left: .5 } }, croppedSamples(128)],
    ["letterbox", { opacity: .5, crop: { left: -.5, right: -.5 } }, letterboxSamples],
  ]) try {
    const program = structuredClone(pairBase);
    program.assets = [backgroundAsset];
    program.pages[0].background = { type: "image", asset: backgroundAsset.id, fit: "stretch", ...fill };
    program.pages[0].elements = [
      { id: "background-control", type: "shape", frame: { x: 100, y: 100, width: 60, height: 40 },
        geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } },
      { id: "shared-image", type: "image", asset: backgroundAsset.id, fit: "stretch", frame: { x: 200, y: 100, width: 60, height: 20 } },
    ];
    const input = { ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)), assets: [{ ...backgroundAsset, data: backgroundData }] };
    const inputHash = sha256(input.program), assetHash = sha256(backgroundData);
    const authored = await compilePpjWorkspace(input, { includePreviewScene: true });
    const originalPaint = await backgroundImagePixels(`${name}-authored`, authored, fill, samples);
    await assertProductionEntry(`background-image-${name}-authored`, input, authored, originalPaint);
    const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
    await writeFile(path.join(artifacts, `background-image-${name}-source.pptx`), source, { flag: "wx" });
    const projection = await projectPptxToPpj(source, { sourceUri: `background-image-${name}.pptx`, assetRootUri: "assets" });
    const bound = { source, assets: projection.assets, program: projection.programJson }, projectionHash = sha256(projection.programJson);
    const noop = await compilePpjWorkspace(bound, { includePreviewScene: true });
    assert.deepEqual(noop.file, source);
    const sourcePaint = await backgroundImagePixels(`${name}-source`, noop, fill, samples);
    await assertProductionEntry(`background-image-${name}-source`, bound, noop, sourcePaint);
    const editEvidence = {};
    if (name === "crop") for (const [edit, expected, editSamples] of [
      ["crop", { opacity: .5, crop: { left: -.5, right: -.5 } }, letterboxSamples],
      ["zero", { opacity: 0, crop: { left: .5 } }, [[.1, transparent], [.5, transparent], [.9, transparent]]],
      ["delete-opacity", { crop: { left: .5 } }, croppedSamples(255)],
      ["delete-crop", { opacity: .5 }, stretchSamples(128)],
      ["delete-background", null, [[.1, [255, 255, 255, 255]], [.5, [255, 255, 255, 255]], [.9, [255, 255, 255, 255]]]],
    ]) try {
      const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8"));
      if (expected === null) delete request.pages[0].background;
      else {
        const background = request.pages[0].background;
        delete background.crop; delete background.opacity;
        Object.assign(background, expected);
      }
      const edited = { ...bound, program: Buffer.from(JSON.stringify(request)) }, requestHash = sha256(edited.program);
      const requestFile = `background-image-source-${edit}.ppj`;
      await writeFile(path.join(artifacts, requestFile), edited.program, { flag: "wx" });
      Object.assign(editEvidence, { requestFile, requestSha256: requestHash, sourceSha256: sourceHash, stage: "compile" });
      const candidate = await compilePpjWorkspace(edited, { includePreviewScene: true });
      editEvidence.stage = "paint/publication";
      const painted = await backgroundImagePixels(`source-${edit}`, candidate, expected, editSamples);
      await assertProductionEntry(`background-image-source-${edit}`, edited, candidate, painted);
      editEvidence.stage = "reprojection/preservation";
      const fresh = await projectPptxToPpj(candidate.file, { sourceUri: `background-image-${edit}.pptx`, assetRootUri: "assets" });
      const observed = JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0].background;
      if (expected === null) assert.equal(observed, undefined);
      else {
        assert.equal(observed.opacity, expected.opacity);
        if (!expected.crop) assert.equal(observed.crop, undefined);
        else for (const side of ["left", "top", "right", "bottom"]) assert.equal(observed.crop[side] ?? 0, expected.crop[side] ?? 0);
      }
      const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file), changed = [];
      assert.deepEqual(Object.keys(newZip.files).sort(), Object.keys(oldZip.files).sort());
      for (const part of Object.keys(oldZip.files)) if (!oldZip.files[part].dir &&
        !Buffer.from(await oldZip.file(part).async("uint8array")).equals(Buffer.from(await newZip.file(part).async("uint8array")))) changed.push(part);
      assert.deepEqual(changed, ["ppt/slides/slide1.xml"], "shared image media and relationships must stay unchanged");
      if (expected === null) assert.ok(!/<p:bg[ >]/.test(await newZip.file("ppt/slides/slide1.xml").async("string")));
      assert.equal(sha256(edited.program), requestHash);
      Object.assign(backgroundImageCases.at(-1), { reprojection: true, changedParts: changed, sourceSha256: sourceHash, requestSha256: requestHash });
    } catch (error) {
      backgroundImageFailures.push({ name: `${name}/${edit}`, ...editEvidence, code: error.code, message: error.message });
    }
    if (name === "stretch") {
      const tiled = structuredClone(program); tiled.pages[0].background.fit = "tile";
      const tiledInput = { ...input, program: Buffer.from(JSON.stringify(tiled)) };
      const tileAuthored = await compilePpjWorkspace(tiledInput, { includePreviewScene: true });
      const tileSource = await withoutAuthoredSnapshot(tileAuthored.file), tileHash = sha256(tileSource);
      const tileProjection = await projectPptxToPpj(tileSource, { sourceUri: "background-tile.pptx", assetRootUri: "assets" });
      const tileBound = { source: tileSource, program: tileProjection.programJson, assets: tileProjection.assets };
      const tileNoop = await compilePpjWorkspace(tileBound, { includePreviewScene: true });
      assert.deepEqual(tileNoop.file, tileSource);
      for (const [origin, receipt, workspace] of [["authored", tileAuthored, tiledInput], ["source", tileNoop, tileBound]]) {
        const painted = await savePaint(`background-image-tile-${origin}`, receipt);
        assert.equal(painted.reliability.status, "failed");
        assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.image-tile" && d.scenePath.endsWith(".background.imagePaint.mode")));
        assert.match(painted.pages[0].svg, /data-officekit-background="unavailable"/);
        assert.doesNotMatch(painted.pages[0].svg, /data-officekit-background="image"/);
        assert.match(painted.pages[0].svg, /data-officekit-native-id=/, "foreground survives unsupported background mode");
        const inputPath = path.join(artifacts, `background-image-tile-${origin}.ppj`);
        const outputDir = path.join(artifacts, `entry-background-image-tile-${origin}`);
        await writeFile(inputPath, workspace.program, { flag: "wx" });
        let failure;
        await assert.rejects(() => renderPpjToSvg(inputPath, { outputDir,
          load: async () => ({ ...workspace, path: inputPath }),
        }), error => { failure = error; return error.code === "preview.output.incomplete"; });
        const persisted = JSON.parse(await readFile(path.join(outputDir, "render.json"), "utf8"));
        assert.deepEqual(JSON.parse(JSON.stringify(failure.receipt)), persisted);
        assert.equal(persisted.ok, false); assert.equal(persisted.output.status, "incomplete");
        assert.equal(persisted.reliability.status, "failed");
        assert.equal(persisted.scene.sha256, receipt.previewScene.sha256);
        assert.equal(persisted.scene.candidateSha256, sha256(receipt.file));
        assert.deepEqual(persisted.failures.map(f => [f.stage, f.message]).sort(),
          [["preview", "preview.scene.paint.background"], ["preview", "preview.scene.paint.image-tile"]].sort());
        assert.ok(persisted.diagnostics.some(d => d.path === "$.pages[0].background.fit"), "original input checks remain mandatory on failed publication");
        for (const artifact of persisted.artifacts)
          assert.equal(sha256(await readFile(path.join(outputDir, artifact.file))), artifact.sha256);
        const png = await readFile(path.join(outputDir, persisted.pages[0].png));
        const region = { left: 0, top: 24, width: persisted.canvas.width, height: persisted.canvas.height - 24 };
        assert.deepEqual(await sharp(png).extract(region).raw().toBuffer(), await sharp(Buffer.from(painted.pages[0].svg)).extract(region).raw().toBuffer());
        assert.deepEqual([...await sharp(png).extract({ left: 1, top: 1, width: 1, height: 1 }).removeAlpha().raw().toBuffer()], [153, 27, 27]);
        assert.deepEqual(await readFile(inputPath), Buffer.from(workspace.program));
        backgroundImageRejections.push({ origin, status: "unavailable", noInventedStretch: true, sourcePreserved: true,
          publication: "incomplete", returnedPersistedEqual: true, retainedArtifactHashes: true, contentAndRedWarningPixels: true });
      }
      assert.equal(sha256(tileSource), tileHash);
    }
    assert.equal(sha256(source), sourceHash); assert.equal(sha256(projection.programJson), projectionHash);
    assert.equal(sha256(input.program), inputHash); assert.equal(sha256(backgroundData), assetHash);
  } catch (error) {
    backgroundImageFailures.push({ name, code: error.code, message: error.message });
    console.error(`Background image ${name} failed: ${error.message}`);
  }
  const shapeBlue = [0, 0, 255, 255], shapeGreen = [0, 128, 127, 255], shapeRed = [128, 0, 127, 255];
  const shapeGeometryReason = "preview.fact.shape-geometry-omitted";
  async function shapeImageProduction(name, input, compiled, painted) {
    const published = await assertProductionEntry(`shape-image-${name}`, input, compiled, painted);
    assert.ok(!published.diagnostics.some(d => d.reason === shapeGeometryReason));
    assert.equal(published.reliability.status, "requires-review", "mapped geometry does not establish complete text/fill fidelity");
    return published;
  }
  const shapeFillGeometry = preset => preset === "custom" ? {
    kind: "custom", viewBox: { x: 0, y: 0, width: 100, height: 100 }, paths: [{ fill: true, stroke: true,
      commands: [{ op: "moveTo", x: 50, y: 0 }, { op: "lineTo", x: 100, y: 50 },
        { op: "lineTo", x: 50, y: 100 }, { op: "lineTo", x: 0, y: 50 }, { op: "close" }] }],
  } : { kind: "preset", preset };
  async function shapeImagePixels(name, receipt, expected, samples) {
    const painted = await savePaint(`shape-image-${name}`, receipt), view = createPpjSceneView(receipt);
    const node = view.pages[0].nodes.find(n => n.kind === "shape"), shape = node.native;
    const fill = shape.imageFill;
    assert.equal(shape.lineRgb, "FF00FF");
    assert.equal(shape.lineWidthEmu, 4n * 12700n);
    if (expected === null) {
      assert.equal(fill, undefined);
      assert.doesNotMatch(painted.pages[0].svg, /data-officekit-shape-image="true"/);
    } else {
      assert.equal(fill.mode, 1);
      assert.equal(fill.opacityThousandthPercent, expected.opacity === undefined ? undefined : expected.opacity * 100000);
      if (!expected.crop) assert.equal(fill.crop, undefined);
      else for (const side of ["left", "top", "right", "bottom"])
        assert.equal(fill.crop[`${side}ThousandthPercent`], (expected.crop[side] ?? 0) * 100000);
      assert.deepEqual(Buffer.from(view.asset(fill.assetId).data), backgroundData);
      assert.match(painted.pages[0].svg, /data-officekit-shape-image="true"/);
    }
    const raster = await sharp(Buffer.from(painted.pages[0].svg)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
    const pixel = (x, y) => [...raster.data.subarray((y * raster.info.width + x) * 4, (y * raster.info.width + x) * 4 + 4)];
    for (const [x, y, rgba] of samples) assert.deepEqual(pixel(x, y), rgba, `${name}: clipped image pixel ${x},${y}`);
    assert.deepEqual(pixel(220, 100), [255, 0, 255, 255], "shape outline retains independent full opacity");
    let blackText = 0;
    for (let y = 105; y < 150; y++) for (let x = 105; x < 335; x++) if (pixel(x, y).slice(0, 3).every(v => v === 0)) blackText++;
    assert.ok(blackText > 50, `${name}: original text remains visible above the fill`);
    assert.match(painted.pages[0].svg, /HHHH/);
    assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.shape-image" || d.scenePath?.includes(".imageFill")));
    const canonical = JSON.parse(new TextDecoder().decode(receipt.programJson));
    const options = { rendererProfile: "native-scene-svg", sceneReceipt: receipt, scenePaint: painted };
    const legacy = assessPpjPreviewInput(canonical), mapped = assessPpjPreviewInput(canonical, options);
    assert.ok(legacy.diagnostics.some(d => d.reason === shapeGeometryReason));
    assert.ok(!mapped.diagnostics.some(d => d.reason === shapeGeometryReason));
    assert.ok(painted.shapeGeometryScenePaths.includes(node.scenePath));
    const missing = assessPpjPreviewInput(canonical, { ...options, scenePaint: { ...painted, shapeGeometryScenePaths: [] } });
    assert.ok(missing.diagnostics.some(d => d.reason === shapeGeometryReason));
    assert.deepEqual(mapped.diagnostics, missing.diagnostics.filter(d => d.reason !== shapeGeometryReason), "no unrelated input rule is relaxed");
    const combined = paintPpjSceneSvg(receipt, { assessInput: true });
    assert.ok(!combined.diagnostics.some(d => d.reason === shapeGeometryReason));
    for (const d of painted.diagnostics)
      assert.ok(combined.diagnostics.some(a => a.scenePath === d.scenePath && a.reason === d.reason && a.status === d.status));
    const registry = JSON.parse(await readFile("src/ppj/capability-registry.json", "utf8"));
    delete registry.previewScene.factualMappings.shapeGeometry;
    assert.throws(() => assessPpjPreviewInput(canonical, { ...options, registry }), /shape geometry mapping/);
    return { painted, record: { name, nativeAndClippedRgbaPixels: true, independentOutlineAndText: true,
      geometryErrorRetired: true, missingCaptureRetainsError: true, missingRegistryRejects: true, unrelatedRulesRetained: true,
      candidateSha256: sha256(receipt.file), sceneSha256: receipt.previewScene.sha256 } };
  }
  for (const preset of ["rect", "roundRect", "ellipse", "diamond", "custom"]) {
    const evidence = { preset, stage: "author compile" };
    try {
      const program = structuredClone(pairBase), fill = { type: "image", asset: backgroundAsset.id, fit: "stretch", opacity: .5, crop: { left: .5 } };
      program.assets = [backgroundAsset]; program.pages[0].background = { type: "solid", color: "#0000FF" };
      program.pages[0].elements = [{ id: "image-filled-shape", type: "shape", frame: { x: 100, y: 100, width: 240, height: 120 },
        geometry: shapeFillGeometry(preset), style: { fill, stroke: { color: "#FF00FF", width: 4 } },
        text: { paragraphs: [{ runs: [{ text: "HHHH", style: { size: 20, color: "#000000" } }] }] } }];
      const input = { ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)), assets: [{ ...backgroundAsset, data: backgroundData }] };
      const inputHash = sha256(input.program), assetHash = sha256(backgroundData);
      const authored = await compilePpjWorkspace(input, { includePreviewScene: true });
      evidence.stage = "author paint";
      const samples = [[124, 160, shapeBlue], [220, 160, shapeGreen], [316, 160, shapeGreen],
        // Stay outside the curved pen's antialiased fringe, while remaining
        // strictly inside the rect's fill, not its 4pt outline.
        [336, 103, preset === "rect" ? shapeGreen : shapeBlue]];
      const original = await shapeImagePixels(`${preset}-authored`, authored, fill, samples);
      await shapeImageProduction(`${preset}-authored`, input, authored, original.painted);
      shapeImageCases.push(original.record);
      const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source);
      const sourceFile = `shape-image-${preset}-source.pptx`;
      await writeFile(path.join(artifacts, sourceFile), source, { flag: "wx" });
      evidence.stage = "source projection/no-op";
      Object.assign(evidence, { sourceFile, sourceSha256: sourceHash });
      const projection = await projectPptxToPpj(source, { sourceUri: sourceFile, assetRootUri: "assets" });
      const bound = { source, assets: projection.assets, program: projection.programJson }, projectionHash = sha256(projection.programJson);
      const noop = await compilePpjWorkspace(bound, { includePreviewScene: true });
      assert.deepEqual(noop.file, source);
      const originalSource = await shapeImagePixels(`${preset}-source`, noop, fill, samples);
      await shapeImageProduction(`${preset}-source`, bound, noop, originalSource.painted);
      shapeImageCases.push(originalSource.record);
      if (preset === "rect") for (const [edit, expected, editSamples] of [
        ["crop", { opacity: .5, crop: { left: -.5, right: -.5 } }, [[124, 160, shapeBlue], [170, 160, shapeRed], [220, 160, shapeBlue], [270, 160, shapeGreen], [316, 160, shapeBlue]]],
        ["zero", { opacity: 0, crop: { left: .5 } }, [[124, 160, shapeBlue], [220, 160, shapeBlue], [316, 160, shapeBlue]]],
        ["delete-opacity", { crop: { left: .5 } }, [[124, 160, shapeBlue], [220, 160, [0, 255, 0, 255]], [316, 160, [0, 255, 0, 255]]]],
        ["delete-crop", { opacity: .5 }, [[124, 160, shapeRed], [220, 160, shapeBlue], [316, 160, shapeGreen]]],
        ["flip", { opacity: .5, crop: { left: .5 } }, [[124, 160, shapeGreen], [220, 160, shapeGreen], [316, 160, shapeBlue]]],
        ["delete-fill", null, [[124, 160, shapeBlue], [220, 160, shapeBlue], [316, 160, shapeBlue]]],
      ]) {
        const request = JSON.parse(Buffer.from(projection.programJson).toString("utf8"));
        const target = request.pages[0].elements.find(e => e.type === "shape");
        if (expected === null) delete target.style.fill;
        else {
          delete target.style.fill.crop; delete target.style.fill.opacity;
          Object.assign(target.style.fill, expected);
          if (edit === "flip") target.frame.flipH = true;
        }
        const edited = { ...bound, program: Buffer.from(JSON.stringify(request)) }, requestHash = sha256(edited.program);
        const requestFile = `shape-image-source-${edit}.ppj`;
        await writeFile(path.join(artifacts, requestFile), edited.program, { flag: "wx" });
        const editEvidence = { preset, edit, sourceFile, sourceSha256: sourceHash, requestFile, requestSha256: requestHash, stage: "compile" };
        try {
          const candidate = await compilePpjWorkspace(edited, { includePreviewScene: true });
          const candidateFile = `shape-image-source-${edit}.pptx`;
          await writeFile(path.join(artifacts, candidateFile), candidate.file, { flag: "wx" });
          editEvidence.stage = "paint/publication";
          const result = await shapeImagePixels(`source-${edit}`, candidate, expected, editSamples);
          if (edit === "flip") assert.equal(createPpjSceneView(candidate).pages[0].nodes[0].native.transform.flipHorizontal, true);
          await shapeImageProduction(`source-${edit}`, edited, candidate, result.painted);
          editEvidence.stage = "reprojection/preservation";
          const fresh = await projectPptxToPpj(candidate.file, { sourceUri: `shape-image-${edit}.pptx`, assetRootUri: "assets" });
          const observed = JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0].elements.find(e => e.type === "shape");
          if (expected === null) assert.equal(observed.style?.fill, undefined);
          else {
            assert.equal(observed.style.fill.opacity, expected.opacity);
            if (!expected.crop) assert.equal(observed.style.fill.crop, undefined);
            else for (const side of ["left", "top", "right", "bottom"])
              assert.equal(observed.style.fill.crop[side] ?? 0, expected.crop[side] ?? 0);
          }
          if (edit === "flip") assert.equal(observed.frame.flipH, true);
          const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file), changed = [];
          const removed = Object.keys(oldZip.files).filter(part => !newZip.files[part]);
          assert.deepEqual(removed, edit === "delete-fill" ? ["ppt/media/image.png"] : []);
          assert.deepEqual(Object.keys(newZip.files).sort(), Object.keys(oldZip.files).filter(part => !removed.includes(part)).sort());
          for (const part of Object.keys(oldZip.files)) if (!oldZip.files[part].dir && newZip.file(part) &&
            !Buffer.from(await oldZip.file(part).async("uint8array")).equals(Buffer.from(await newZip.file(part).async("uint8array")))) changed.push(part);
          if (edit === "delete-fill") {
            // This sole-use media part belongs to the removed fill. Only that
            // part and relationship may disappear; all non-target ZIP bytes
            // are still compared above. Shared-media deletion is a separate
            // retained background regression, not bypassed by this fixture.
            assert.deepEqual(Buffer.from(await oldZip.file("ppt/media/image.png").async("uint8array")), backgroundData);
            const rels = "ppt/slides/_rels/slide1.xml.rels", before = await oldZip.file(rels).async("string");
            const owned = [...before.matchAll(/<Relationship\b(?=[^>]*Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/image")(?=[^>]*Target="\/ppt\/media\/image.png")[^>]*\/>/g)];
            assert.equal(owned.length, 1);
            assert.equal(await newZip.file(rels).async("string"), before.replace(owned[0][0], ""));
            assert.doesNotMatch(await newZip.file("ppt/slides/slide1.xml").async("string"), /<a:blipFill\b/);
            assert.deepEqual(changed.sort(), [rels, "ppt/slides/slide1.xml"].sort());
          } else assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
          shapeImageCases.push({ ...result.record, candidateFile, reprojection: true, changedParts: changed,
            removedParts: removed, sourceSha256: sourceHash, requestSha256: requestHash });
        } catch (error) {
          shapeImageFailures.push({ ...editEvidence, code: error.code, message: error.message });
        }
        assert.equal(sha256(edited.program), requestHash);
        assert.equal(sha256(source), sourceHash);
      }
      assert.equal(sha256(source), sourceHash); assert.equal(sha256(projection.programJson), projectionHash);
      assert.equal(sha256(input.program), inputHash); assert.equal(sha256(backgroundData), assetHash);
    } catch (error) {
      shapeImageFailures.push({ ...evidence, code: error.code, message: error.message });
      console.error(`Shape image ${preset} failed at ${evidence.stage}: ${error.message}`);
    }
  }
  // Independent literal paths, not output sampled from the polygon painter.
  // The 240x120 frame makes short-side adjustment scaling distinguishable
  // from the incorrect width-based construction.
  const polygonReferencePoints = (preset, adjustment) => {
    const offset = { undefined: preset === "chevron" ? 60 : preset === "triangle" ? 120 : 30,
      75000: preset === "triangle" ? 180 : 90, 25000: preset === "triangle" ? 60 : 30, 0: 0 }[String(adjustment)];
    switch (preset) {
      case "triangle": return [[0,120],[offset,0],[240,120]];
      case "rtTriangle": return [[0,120],[0,0],[240,120]];
      case "trapezoid": return [[0,120],[offset,0],[240-offset,0],[240,120]];
      case "parallelogram": return [[0,120],[offset,0],[240,0],[240-offset,120]];
      case "chevron": return [[0,0],[240-offset,0],[240,60],[240-offset,120],[0,120],[offset,60]];
    }
  };
  const polygonReferenceTextRectangle = (preset, adjustment) => {
    const variant = String(adjustment);
    let edges;
    if (preset === "triangle") edges = { undefined: [60,60,180,120], 75000: [90,60,210,120], 25000: [30,60,150,120], 0: [0,60,120,120] }[variant];
    if (preset === "rtTriangle") edges = [20,70,140,110];
    if (preset === "trapezoid") edges = { undefined: [20,10,220,120], 75000: [60,30,180,120], 25000: [20,10,220,120], 0: [0,0,240,120] }[variant];
    if (preset === "parallelogram") edges = { undefined: [32.5,16.25,207.5,103.75], 75000: [57.5,28.75,182.5,91.25], 25000: [32.5,16.25,207.5,103.75], 0: [20,10,220,110] }[variant];
    if (preset === "chevron") edges = { undefined: [60,0,180,120], 75000: [90,0,150,120], 25000: [30,0,210,120], 0: [0,0,240,120] }[variant];
    assert.ok(edges);
    return Object.fromEntries(["left", "top", "right", "bottom"].map((k,i) => [k,edges[i]]));
  };
  const exceptFirstPresetAdjustments = xml => {
    // Semantic image edits may add a redundant relationship namespace at the
    // slide root. Prove every r-qualified use already declares that same URI
    // locally before normalizing only this exact root declaration.
    const relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    for (const [tag] of xml.matchAll(/<[^>]+>/g)) if (/\sr:[\w.-]+="/.test(tag))
      assert.equal(xmlAttributes(tag)["xmlns:r"], relationships, "relationship prefix must already have the same local binding");
    xml = xml.replace(/<p:sld\b[^>]*>/, tag => tag.replace(` xmlns:r="${relationships}"`, ""));
    const drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const rootDrawing = xmlAttributes(xml.match(/<p:sld\b[^>]*>/)[0])["xmlns:a"];
    xml = xml.replace(/<p:blipFill\b[^>]*>[\s\S]*?<\/p:blipFill>/, fill =>
      fill.replace(/<a:srcRect\b[^>]*\/>/, tag => {
        const local = xmlAttributes(tag)["xmlns:a"];
        assert.ok(local === drawing || local === undefined && rootDrawing === drawing,
          "picture crop must have the same local or root DrawingML binding");
        return tag.replace(` xmlns:a="${drawing}"`, "");
      }));
    let geometries = 0, lists = 0;
    const masked = xml.replace(/<a:prstGeom\b[^>]*>[\s\S]*?<\/a:prstGeom>/, geometry => {
      geometries++;
      return geometry.replace(/<a:avLst\b[^>]*(?:\/>|>[\s\S]*?<\/a:avLst>)/, () => { lists++; return "<a:avLst/>"; });
    });
    assert.equal(geometries, 1); assert.equal(lists, 1);
    return orderedXml(masked);
  };
  function polygonProgram(preset, consumer, adjustment, reference = false) {
    const program = structuredClone(pairBase), points = polygonReferencePoints(preset, adjustment);
    const geometry = reference ? { kind: "custom", viewBox: { x: 0, y: 0, width: 240, height: 120 }, paths: [{
      fill: true, stroke: true, commands: [...points.map(([x,y],i) => ({ op: i ? "lineTo" : "moveTo", x,y })), { op: "close" }],
    }], ...(consumer === "image" ? {} : { textRectangle: polygonReferenceTextRectangle(preset, adjustment) }) }
      : { kind: "preset", preset, ...(adjustment === undefined ? {} : { adjustments: [adjustment] }) };
    const frame = { x: 100, y: 100, width: 240, height: 120,
      ...(preset === "chevron" ? { rotation: 15, flipH: true } : {}) };
    const imageFill = { type: "image", asset: backgroundAsset.id, fit: "stretch", opacity: .5, crop: { left: -.5, right: -.5 } };
    const stroke = { color: "#FF00FF", width: 4 };
    const element = consumer === "image" ? { id: "polygon", type: "image", frame, mask: geometry,
      asset: backgroundAsset.id, fit: "stretch", crop: imageFill.crop, opacity: .5, border: stroke }
      : { id: "polygon", type: "shape", frame, geometry,
        style: { fill: consumer === "shape-image" ? imageFill : { type: "solid", color: "#008800" }, stroke },
        text: { paragraphs: [{ runs: [{ text: "F0", style: { size: 16, color: "#000000" } }] }] } };
    program.assets = consumer === "shape" ? [] : [backgroundAsset];
    program.pages[0].background = { type: "solid", color: "#0000FF" };
    program.pages[0].elements = [element, { id: "unrelated", type: "shape", frame: { x: 400, y: 300, width: 50, height: 50 },
      geometry: { kind: "preset", preset: "rect" }, style: { fill: { type: "solid", color: "#CC5500" } } }];
    return { ...sourceWorkspace, program: Buffer.from(JSON.stringify(program)), assets: consumer === "shape" ? [] : [{ ...backgroundAsset, data: backgroundData }] };
  }
  async function polygonPixels(name, receipt, input, preset, consumer, adjustment) {
    const view = createPpjSceneView(receipt), target = view.pages[0].nodes[0], native = target.native;
    assert.equal(consumer === "image" ? native.maskPreset : native.geometry, preset);
    assert.deepEqual(consumer === "image" ? native.maskPresetAdjustments : native.presetAdjustments, adjustment === undefined ? [] : [adjustment]);
    const painted = await savePaint(`polygon-${name}`, receipt);
    assert.ok(!painted.diagnostics.some(d => /preview.scene.paint.(preset|preset-adjustments|image-mask|shape-image)$/.test(d.reason)));
    const referenceInput = polygonProgram(preset, consumer, adjustment, true);
    const reference = await compilePpjWorkspace(referenceInput, { includePreviewScene: true });
    const referencePaint = paintPpjSceneSvg(reference), region = { left: 0, top: 24, width: painted.canvas.width, height: painted.canvas.height - 24 };
    const pixels = svg => sharp(Buffer.from(svg)).extract(region).ensureAlpha().raw().toBuffer();
    assert.deepEqual(await pixels(painted.pages[0].svg), await pixels(referencePaint.pages[0].svg), `${name}: independent literal contour, fill/alpha/border/text and outer transform`);
    const publicResult = await assertProductionEntry(`polygon-${name}`, input, receipt, painted);
    assert.ok(!publicResult.diagnostics.some(d => d.reason === shapeGeometryReason));
    const candidateFile = `polygon-${name}.pptx`;
    await writeFile(path.join(artifacts, candidateFile), receipt.file, { flag: "wx" });
    return { name, preset, consumer, adjustment, candidateFile, candidateSha256: sha256(receipt.file),
      sceneSha256: receipt.previewScene.sha256, nativeAndLiteralContourRaster: true, publicReliability: publicResult.reliability.status };
  }
  for (const preset of ["triangle", "rtTriangle", "trapezoid", "parallelogram", "chevron"])
    for (const consumer of ["shape", "shape-image", "image"]) {
      const name = `${preset}-${consumer}`, adjustment = preset === "rtTriangle" ? undefined : 75000;
      let stage = "author";
      try {
        const input = polygonProgram(preset, consumer, adjustment), inputHash = sha256(input.program);
        const authored = await compilePpjWorkspace(input, { includePreviewScene: true });
        polygonPresetCases.push(await polygonPixels(`${name}-author`, authored, input, preset, consumer, adjustment));
        const source = await withoutAuthoredSnapshot(authored.file), sourceHash = sha256(source), sourceFile = `polygon-${name}-source.pptx`;
        await writeFile(path.join(artifacts, sourceFile), source, { flag: "wx" });
        stage = "source-noop";
        const projected = await projectPptxToPpj(source, { sourceUri: sourceFile, assetRootUri: "assets" });
        const sourceInput = { source, assets: projected.assets, program: projected.programJson };
        const noop = await compilePpjWorkspace(sourceInput, { includePreviewScene: true });
        assert.deepEqual(noop.file, source);
        polygonPresetCases.push({ ...await polygonPixels(`${name}-noop`, noop, sourceInput, preset, consumer, adjustment), sourceFile, sourceSha256: sourceHash });
        if (preset !== "rtTriangle") for (const value of [25000, 0, undefined]) {
          stage = `source-${value ?? "delete"}`;
          // Start each change from the original bytes, not the prior candidate.
          const freshSource = await projectPptxToPpj(source, { sourceUri: sourceFile, assetRootUri: "assets" });
          const program = JSON.parse(Buffer.from(freshSource.programJson).toString("utf8")), geometry = program.pages[0].elements[0][consumer === "image" ? "mask" : "geometry"];
          if (value === undefined) delete geometry.adjustments; else geometry.adjustments = [value];
          const edit = { source, assets: freshSource.assets, program: Buffer.from(JSON.stringify(program)) };
          const requestFile = `polygon-${name}-${stage}.ppj`, requestHash = sha256(edit.program);
          await writeFile(path.join(artifacts, requestFile), edit.program, { flag: "wx" });
          const candidate = await compilePpjWorkspace(edit, { includePreviewScene: true });
          const record = await polygonPixels(`${name}-${stage}`, candidate, edit, preset, consumer, value);
          const fresh = await projectPptxToPpj(candidate.file, { sourceUri: record.candidateFile, assetRootUri: "assets" });
          const target = JSON.parse(Buffer.from(fresh.programJson).toString("utf8")).pages[0].elements[0];
          assert.deepEqual(target[consumer === "image" ? "mask" : "geometry"].adjustments, value === undefined ? undefined : [value]);
          const before = createPpjSceneView(noop).pages[0].nodes[0].native, after = createPpjSceneView(candidate).pages[0].nodes[0].native;
          const field = consumer === "image" ? "maskPresetAdjustments" : "presetAdjustments";
          assert.deepEqual({ ...after, [field]: [] }, { ...before, [field]: [] }, "all non-adjustment native fields remain exact");
          const oldZip = await JSZip.loadAsync(source), newZip = await JSZip.loadAsync(candidate.file), changed = [];
          assert.deepEqual(Object.keys(newZip.files).sort(), Object.keys(oldZip.files).sort());
          for (const file of Object.keys(oldZip.files)) if (!oldZip.files[file].dir &&
            !Buffer.from(await oldZip.file(file).async("uint8array")).equals(Buffer.from(await newZip.file(file).async("uint8array")))) changed.push(file);
          assert.deepEqual(changed, ["ppt/slides/slide1.xml"]);
          const oldXml = await oldZip.file("ppt/slides/slide1.xml").async("string"), newXml = await newZip.file("ppt/slides/slide1.xml").async("string");
          assert.equal(exceptFirstPresetAdjustments(newXml), exceptFirstPresetAdjustments(oldXml), "only the target preset adjustment list may change within the slide");
          assert.notEqual(exceptFirstPresetAdjustments(newXml.replace('val="CC5500"', 'val="112233"')),
            exceptFirstPresetAdjustments(oldXml), "unrelated sibling paint mutations remain detected");
          if (consumer !== "shape") assert.throws(() => exceptFirstPresetAdjustments(newXml.replaceAll(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "urn:changed-relationship-namespace")), /same local binding/);
          const reprojectionFile = `polygon-${name}-${stage}-reprojected.ppj`;
          await writeFile(path.join(artifacts, reprojectionFile), fresh.programJson, { flag: "wx" });
          assert.equal(sha256(edit.program), requestHash); assert.equal(sha256(source), sourceHash);
          polygonPresetCases.push({ ...record, sourceFile, sourceSha256: sourceHash, requestFile, requestSha256: requestHash,
            reprojectionFile, reprojectionSha256: sha256(fresh.programJson), changedParts: changed, nonAdjustmentStatePreserved: true });
        }
        assert.equal(sha256(input.program), inputHash); assert.equal(sha256(source), sourceHash); assert.equal(sha256(backgroundData), backgroundAsset.sha256);
      } catch (error) {
        polygonPresetFailures.push({ preset, consumer, stage, code: error.code, message: error.message });
        console.error(`Polygon ${name}/${stage} failed: ${error.message}`);
      }
    }
  cropProgram.assets = [cropAsset];
  cropProgram.pages[0].elements = [{ id: "crop-picture", type: "image", asset: cropAsset.id, frame: { x: 100, y: 100, width: 100, height: 100 }, fit: "stretch", crop: { left: 0.5 } }];
  cropProgram.pages[0].elements.unshift({ ...structuredClone(primitive), id: "crop-background", geometry: { kind: "preset", preset: "rect" }, frame: { x: 100, y: 100, width: 100, height: 100 }, text: undefined, style: { fill: { type: "solid", color: "#0000FF" } } });
  const cropAuthored = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(cropProgram)), assets: [{ ...cropAsset, data: cropData }] }, { includePreviewScene: true });
  for (const cropped of [false, true]) {
    const tiledProgram = structuredClone(cropProgram), picture = tiledProgram.pages[0].elements.find(e => e.type === "image");
    picture.fit = "tile";
    if (!cropped) delete picture.crop;
    const authoredTile = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(tiledProgram)), assets: [{ ...cropAsset, data: cropData }] }, { includePreviewScene: true });
    const tileSource = await withoutAuthoredSnapshot(authoredTile.file), originalTileHash = sha256(tileSource);
    const tileProjected = await projectPptxToPpj(tileSource, { sourceUri: "tile.pptx", assetRootUri: "assets" });
    const sourceTile = await compilePpjWorkspace({ program: tileProjected.programJson, source: tileSource, assets: tileProjected.assets }, { includePreviewScene: true });
    assert.deepEqual(sourceTile.file, tileSource);
    for (const [kind, receipt] of [["authored", authoredTile], ["source", sourceTile]]) {
      const node = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "image");
      assert.equal(node.native.tiled, true);
      const painted = await savePaint(`tile-${cropped}-${kind}`, receipt);
      assert.equal(painted.reliability.status, "failed");
      assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.image-tile" && d.scenePath?.endsWith("image.tiled")));
      assert.match(painted.pages[0].svg, /Tiled image unavailable/);
      assert.doesNotMatch(painted.pages[0].svg, /<image /, "tile cannot silently become a stretched picture");
    }
    assert.equal(sha256(tileSource), originalTileHash);
  }
  async function cropPixels(receipt, name, samples) {
    const painted = await savePaint(name, receipt);
    for (const [x, rgb, y = 150] of samples) {
      const pixel = await sharp(Buffer.from(painted.pages[0].svg)).extract({ left: x, top: y, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
      assert.deepEqual([...pixel], rgb);
    }
    assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.image-crop"));
  }
  await cropPixels(cropAuthored, "crop-authored", [[120, [0, 255, 0]], [180, [0, 255, 0]], [210, [255, 255, 255]]]);
  const cropSource = await withoutAuthoredSnapshot(cropAuthored.file), cropOriginal = cropSource.slice();
  const cropProjected = await projectPptxToPpj(cropSource, { sourceUri: "crop.pptx", assetRootUri: "assets" });
  const cropInput = { source: cropSource, assets: cropProjected.assets, program: cropProjected.programJson };
  const cropNoop = await compilePpjWorkspace(cropInput, { includePreviewScene: true });
  assert.deepEqual(cropNoop.file, cropSource);
  await cropPixels(cropNoop, "crop-noop", [[120, [0, 255, 0]]]);
  const cropEdit = JSON.parse(Buffer.from(cropProjected.programJson).toString("utf8"));
  cropEdit.pages[0].elements.find(e => e.type === "image").crop = { left: -0.5, right: -0.5 };
  const cropCandidate = await compilePpjWorkspace({ ...cropInput, program: Buffer.from(JSON.stringify(cropEdit)) }, { includePreviewScene: true });
  await cropPixels(cropCandidate, "crop-letterbox", [[110, [0, 0, 255]], [135, [255, 0, 0]], [165, [0, 255, 0]], [190, [0, 0, 255]]]);
  const cropFresh = await projectPptxToPpj(cropCandidate.file, { sourceUri: "crop-edited.pptx", assetRootUri: "assets" });
  assert.equal(JSON.parse(Buffer.from(cropFresh.programJson).toString("utf8")).pages[0].elements.find(e => e.type === "image").crop.left, -0.5);
  assert.deepEqual(cropSource, cropOriginal);
  const cropOldZip = await JSZip.loadAsync(cropSource), cropNewZip = await JSZip.loadAsync(cropCandidate.file);
  assert.deepEqual(Object.keys(cropNewZip.files).sort(), Object.keys(cropOldZip.files).sort());
  const cropChangedParts = [];
  for (const name of Object.keys(cropOldZip.files)) if (!cropOldZip.files[name].dir &&
    !Buffer.from(await cropOldZip.file(name).async("uint8array")).equals(Buffer.from(await cropNewZip.file(name).async("uint8array")))) cropChangedParts.push(name);
  assert.deepEqual(cropChangedParts, ["ppt/slides/slide1.xml"]);
  const borderedProgram = structuredClone(cropProgram);
  borderedProgram.pages[0].elements.find(e => e.type === "image").border = { color: "#FF00FF", width: 4 };
  const borderedAuthored = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(borderedProgram)), assets: [{ ...cropAsset, data: cropData }] }, { includePreviewScene: true });
  await cropPixels(borderedAuthored, "image-border-authored", [[150, [255, 0, 255], 100], [150, [0, 255, 0], 110]]);
  const borderedEdit = JSON.parse(Buffer.from(cropProjected.programJson).toString("utf8"));
  borderedEdit.pages[0].elements.find(e => e.type === "image").border = { color: "#FF00FF", width: 4 };
  const borderedCandidate = await compilePpjWorkspace({ ...cropInput, program: Buffer.from(JSON.stringify(borderedEdit)) }, { includePreviewScene: true });
  await cropPixels(borderedCandidate, "image-border-edited", [[150, [255, 0, 255], 100], [150, [0, 255, 0], 110]]);
  const borderedFresh = await projectPptxToPpj(borderedCandidate.file, { sourceUri: "border.pptx", assetRootUri: "assets" });
  assert.equal(JSON.parse(Buffer.from(borderedFresh.programJson).toString("utf8")).pages[0].elements.find(e => e.type === "image").border.width, 4);
  const borderZip = await JSZip.loadAsync(borderedCandidate.file);
  assert.deepEqual(Object.keys(borderZip.files).sort(), Object.keys(cropOldZip.files).sort());
  for (const name of Object.keys(cropOldZip.files)) if (!cropOldZip.files[name].dir && name !== "ppt/slides/slide1.xml")
    assert.deepEqual(await borderZip.file(name).async("uint8array"), await cropOldZip.file(name).async("uint8array"));
  assert.deepEqual(cropSource, cropOriginal);
  for (const preset of ["ellipse", "diamond", "custom", "roundRect"]) {
    const maskGeometry = preset === "custom" ? { kind: "custom", viewBox: { x: 0, y: 0, width: 100, height: 100 }, paths: [{ fill: true, stroke: false, commands: [
      { op: "moveTo", x: 50, y: 0 }, { op: "lineTo", x: 100, y: 50 }, { op: "lineTo", x: 50, y: 100 }, { op: "lineTo", x: 0, y: 50 }, { op: "close" },
    ] }] } : { kind: "preset", preset, ...(preset === "roundRect" ? { adjustments: [50000] } : {}) };
    const maskProgram = structuredClone(cropProgram);
    maskProgram.pages[0].elements.find(e => e.type === "image").mask = maskGeometry;
    const maskAuthored = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(maskProgram)), assets: [{ ...cropAsset, data: cropData }] }, { includePreviewScene: true });
    await cropPixels(maskAuthored, `mask-${preset}-authored`, [[150, [0, 255, 0]], [110, [0, 0, 255], 110]]);
    const maskEdit = JSON.parse(Buffer.from(cropProjected.programJson).toString("utf8"));
    maskEdit.pages[0].elements.find(e => e.type === "image").mask = maskGeometry;
    const maskCandidate = await compilePpjWorkspace({ ...cropInput, program: Buffer.from(JSON.stringify(maskEdit)) }, { includePreviewScene: true });
    await cropPixels(maskCandidate, `mask-${preset}-edited`, [[150, [0, 255, 0]], [110, [0, 0, 255], 110]]);
    const maskFresh = await projectPptxToPpj(maskCandidate.file, { sourceUri: `${preset}.pptx`, assetRootUri: "assets" });
    const freshMask = JSON.parse(Buffer.from(maskFresh.programJson).toString("utf8")).pages[0].elements.find(e => e.type === "image").mask;
    if (preset === "custom") {
      assert.equal(freshMask.kind, "custom");
      assert.deepEqual(freshMask.paths[0].commands, maskGeometry.paths[0].commands);
    } else assert.equal(freshMask.preset, preset);
    if (preset === "roundRect") {
      assert.deepEqual(freshMask.adjustments, [50000]);
      const squareEdit = JSON.parse(Buffer.from(maskFresh.programJson).toString("utf8"));
      squareEdit.pages[0].elements.find(e => e.type === "image").mask.adjustments = [0];
      const squareCandidate = await compilePpjWorkspace({ source: maskCandidate.file, assets: maskFresh.assets, program: Buffer.from(JSON.stringify(squareEdit)) }, { includePreviewScene: true });
      await cropPixels(squareCandidate, "mask-roundRect-zero", [[110, [0, 255, 0], 110], [150, [0, 255, 0]]]);
      const squareFresh = await projectPptxToPpj(squareCandidate.file, { sourceUri: "square.pptx", assetRootUri: "assets" });
      assert.deepEqual(JSON.parse(Buffer.from(squareFresh.programJson).toString("utf8")).pages[0].elements.find(e => e.type === "image").mask.adjustments, [0]);
    }
    assert.deepEqual(cropSource, cropOriginal);
    const maskZip = await JSZip.loadAsync(maskCandidate.file);
    assert.deepEqual(Object.keys(maskZip.files).sort(), Object.keys(cropOldZip.files).sort());
    for (const name of Object.keys(cropOldZip.files)) if (!cropOldZip.files[name].dir && name !== "ppt/slides/slide1.xml")
      assert.deepEqual(await maskZip.file(name).async("uint8array"), await cropOldZip.file(name).async("uint8array"));
  }
  literalProgram.pages[0].elements = [structuredClone(literalEdge)];
  const literalResult = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(literalProgram)) }, { includePreviewScene: true });
  assertLiteralEdge(literalResult, await savePaint("literal-connector", literalResult));
  const bendProgram = structuredClone(literalProgram);
  Object.assign(bendProgram.pages[0].elements[0], { connectorType: "elbow", bendAdjustment: 25000 });
  const bendAuthored = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(JSON.stringify(bendProgram)) }, { includePreviewScene: true });
  async function assertBend(receipt, value, name) {
    const edge = createPpjSceneView(receipt).pages[0].nodes.find(n => n.kind === "connector");
    assert.equal(edge.native.bendAdjustment, value);
    const painted = await savePaint(name, receipt), mid = 750 - 240 * value / 100000;
    assert.ok(painted.pages[0].svg.includes(`d="M 750 320 L ${mid} 320 L ${mid} 120 L 510 120"`));
    assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.connector-bend"));
    const pixel = await sharp(Buffer.from(painted.pages[0].svg)).extract({ left: mid, top: 220, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
    assert.deepEqual([...pixel], [17, 68, 119]);
    assertDirectEndpointProfile(receipt, painted);
  }
  await assertBend(bendAuthored, 25000, "bend-authored");
  const bendSource = await withoutAuthoredSnapshot(bendAuthored.file), bendOriginal = bendSource.slice();
  const bendProjected = await projectPptxToPpj(bendSource, { sourceUri: "bend.pptx", assetRootUri: "assets" });
  const bendInput = { source: bendSource, assets: bendProjected.assets, program: bendProjected.programJson };
  const bendNoop = await compilePpjWorkspace(bendInput, { includePreviewScene: true });
  assert.deepEqual(bendNoop.file, bendSource);
  await assertBend(bendNoop, 25000, "bend-noop");
  const bendEdit = JSON.parse(Buffer.from(bendProjected.programJson).toString("utf8"));
  bendEdit.pages[0].elements.find(e => e.type === "connector").bendAdjustment = 75000;
  const bendCandidate = await compilePpjWorkspace({ ...bendInput, program: Buffer.from(JSON.stringify(bendEdit)) }, { includePreviewScene: true });
  await assertBend(bendCandidate, 75000, "bend-edited");
  const bendFresh = await projectPptxToPpj(bendCandidate.file, { sourceUri: "bend-edited.pptx", assetRootUri: "assets" });
  assert.equal(JSON.parse(Buffer.from(bendFresh.programJson).toString("utf8")).pages[0].elements.find(e => e.type === "connector").bendAdjustment, 75000);
  assert.deepEqual(bendSource, bendOriginal);
  const bendOldZip = await JSZip.loadAsync(bendSource), bendNewZip = await JSZip.loadAsync(bendCandidate.file);
  assert.deepEqual(Object.keys(bendNewZip.files).sort(), Object.keys(bendOldZip.files).sort());
  const bendChangedParts = [];
  for (const name of Object.keys(bendOldZip.files)) if (!bendOldZip.files[name].dir &&
    !Buffer.from(await bendOldZip.file(name).async("uint8array")).equals(Buffer.from(await bendNewZip.file(name).async("uint8array")))) bendChangedParts.push(name);
  assert.deepEqual(bendChangedParts, ["ppt/slides/slide1.xml"]);
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
  await assertProductionEntry("source-noop", sourceInput, noop, noopPaint);
  await assertProductionEntry("source-edit", editInput, candidate, candidatePaint);
  const repeatedBenchmark = structuredClone(pairBase);
  repeatedBenchmark.design.canvas.width = 4000;
  repeatedBenchmark.components = structuredClone(nestedFixture.components);
  const repeatedInstance = structuredClone(nestedFixture.instance);
  repeatedInstance.frame.width = 3800;
  repeatedInstance.repeat.items = Array.from({ length: 20 }, (_, i) => ({ key: `sample-${i}`, arguments: {} }));
  repeatedBenchmark.pages[0].elements = [repeatedInstance, nestedFixture.risk];
  const performanceCases = [];
  for (const [name, input] of [["repeated-20", { ...sourceWorkspace, program: Buffer.from(JSON.stringify(repeatedBenchmark)) }],
    ["source-candidate", editInput]]) {
    const inputHash = sha256(input.program), sourceHash = input.source ? sha256(input.source) : null;
    // Warm both paths before measurement. These samples are not cold-start
    // claims, memory peaks, or a full G-16 performance acceptance threshold.
    await compilePpjWorkspace(input);
    await compilePpjWorkspace(input, { includePreviewScene: true });
    const samples = [];
    for (let round = 0; round < 3; round++) {
      let ordinaryHash;
      for (const includePreviewScene of [false, true]) {
        measuredTransport = [];
        const memoryBefore = process.memoryUsage(), start = performance.now();
        let compiled, transport;
        try {
          compiled = await compilePpjWorkspace(input, { includePreviewScene });
          transport = measuredTransport;
        } finally { measuredTransport = undefined; }
        const compileMs = performance.now() - start, memoryAfterCompile = process.memoryUsage();
        assert.equal(transport.length, 1, "one measured compile request, no hidden second pass");
        const candidateHash = sha256(compiled.file);
        if (!includePreviewScene) {
          assert.equal(Object.hasOwn(compiled, "previewScene"), false);
          ordinaryHash = candidateHash;
        } else assert.equal(candidateHash, ordinaryHash, "scene collection cannot change the candidate");
        let svgMs = null, pngMs = null, svgBytes = 0, pngBytes = 0, nodeBindings = 0;
        if (includePreviewScene) {
          nodeBindings = compiled.previewScene.bindings.length;
          if (name === "repeated-20") assert.equal(nodeBindings, 41);
          if (name === "source-candidate") assert.equal(compiled.previewScene.origin, 2);
          const paintStart = performance.now(), painted = paintPpjSceneSvg(compiled);
          svgMs = performance.now() - paintStart;
          svgBytes = painted.pages.reduce((sum, page) => sum + Buffer.byteLength(page.svg), 0);
          const rasterStart = performance.now();
          for (const page of painted.pages) pngBytes += (await sharp(Buffer.from(page.svg)).png().toBuffer()).byteLength;
          pngMs = performance.now() - rasterStart;
        }
        assert.equal(sha256(input.program), inputHash);
        assert.equal(input.source ? sha256(input.source) : null, sourceHash);
        samples.push({ round, includePreviewScene, ...transport[0], compileMs, svgMs, pngMs,
          candidateBytes: compiled.file.byteLength, candidateHash, nodeBindings, svgBytes, pngBytes,
          jsMemorySnapshots: { before: memoryBefore, afterCompile: memoryAfterCompile, afterRaster: process.memoryUsage() } });
      }
    }
    const isolatedNative = [];
    let isolatedCandidate;
    for (const includePreviewScene of [false, true]) {
      const result = await fullWireCompile(input, { includePreviewScene, measureNativeMemory: true });
      if (!includePreviewScene) {
        assert.equal(result.program.previewScene, undefined);
        isolatedCandidate = sha256(result.file);
      } else assert.equal(sha256(result.file), isolatedCandidate);
      assert.equal(sha256(input.program), inputHash);
      assert.equal(input.source ? sha256(input.source) : null, sourceHash);
      isolatedNative.push({ includePreviewScene, candidateHash: sha256(result.file), ...result.nativeMemory });
    }
    let sceneBudgetRecovery;
    if (name === "repeated-20") {
      const client = await startOfficeKitNativeClient({ packageJsonPath, profile: "ppj" });
      try {
        const request = create(CodecRequestSchema, { protocolVersion: 2, operation: 11, family: 2,
          limits: { maxUncompressedBytes: 4096n },
          presentationProgram: { programJson: input.program, includePreviewScene: true, includeNodeMap: true } });
        const invoke = async () => fromBinary(CodecResponseSchema,
          await client.invoke(toBinary(CodecRequestSchema, request)), { recursionLimit: 136 });
        const failed = await invoke();
        assert.equal(failed.ok, false);
        assert.ok(failed.diagnostics.some(d => d.code === "preview_scene_budget_exceeded"), JSON.stringify(failed.diagnostics));
        assert.equal(failed.file.byteLength, 0, "budget failure must not publish a candidate");
        assert.equal(failed.presentationProgram?.previewScene, undefined, "budget failure must not publish a truncated scene");
        request.limits = undefined;
        const recovered = await invoke();
        assert.equal(recovered.ok, true, JSON.stringify(recovered.diagnostics));
        const recoveredScene = readPpjPreviewScene(recovered.presentationProgram, recovered.file);
        assert.equal(recoveredScene.bindings.length, 41);
        assert.equal(sha256(recovered.file), isolatedCandidate);
        assert.equal(sha256(input.program), inputHash);
        sceneBudgetRecovery = { status: "passed", maxUncompressedBytes: 4096, error: "preview_scene_budget_exceeded",
          emptyFailedCandidate: true, noTruncatedScene: true, sameProcessRecovery: true, recoveredBindings: 41 };
      } finally { await client.retire(); }
    }
    performanceCases.push({ name, inputHash, sourceHash, samples, isolatedNative, sceneBudgetRecovery });
  }
  await writeFile(path.join(artifacts, "performance.json"), JSON.stringify({
    scope: "Three warm samples per mode plus separate fresh native processes for scene-off/on RSS/high-water. JS snapshots and native OS residency do not prove retained-object release. No performance threshold claim.",
    node: process.version, platform: process.platform, arch: process.arch, cases: performanceCases,
  }, null, 2), { flag: "wx" });
  for (const [name, input, compiled, painted] of [
    ["source-noop", sourceInput, noop, noopPaint], ["source-edited", editInput, candidate, candidatePaint],
  ]) {
    const outputDir = path.join(artifacts, `published-${name}`);
    const evidence = previewInputEvidence(input, compiled);
    publicationCases.push({ name, input, compiled, painted });
    const publication = await publishPpjPreview(painted, evidence, { outputDir });
    const persisted = JSON.parse(await readFile(path.join(outputDir, "render.json"), "utf8"));
    assert.deepEqual(persisted, JSON.parse(JSON.stringify(publication.receipt)));
    assert.deepEqual(persisted.scene, painted.sceneEvidence);
    assert.equal(persisted.scene.origin, "candidate-import");
    assert.equal(persisted.scene.candidateSha256, sha256(compiled.file));
    assert.equal(persisted.source.sha256, sha256(source));
    assert.equal(persisted.scene.candidateSha256 === persisted.source.sha256, name === "source-noop");
    assert.equal(persisted.ok, true);
    await assertPublishedWarning(persisted, painted);
    for (const artifact of persisted.artifacts)
      assert.equal(sha256(await readFile(path.join(outputDir, artifact.file))), artifact.sha256);
  }
  // Old/no-op scene evidence must never be published as the edited candidate.
  const mismatchDir = path.join(artifacts, "rejected-source-identity");
  await assert.rejects(publishPpjPreview(candidatePaint, previewInputEvidence(sourceInput, noop), {
    outputDir: mismatchDir, loadRaster: async () => { assert.fail("identity failure must precede raster loading"); },
  }), error => error.code === "preview.output.scene" && error.receipt.reliability.status === "failed");
  await assert.rejects(readFile(path.join(mismatchDir, "render.pending.json")), { code: "ENOENT" });
  // Real compiled/painted authored and candidate inputs, with one isolated
  // operational fault per output directory. Nonfailed rasters use real sharp.
  const combinedAssessmentCases = [];
  for (const { name, input, compiled, painted: paintOnly } of publicationCases) {
    const originalCandidate = sha256(compiled.file), originalInput = Buffer.from(compiled.programJson);
    const painted = paintPpjSceneSvg(compiled, { assessInput: true });
    assert.ok(painted.inputAssessment);
    for (const d of painted.inputAssessment.diagnostics)
      assert.ok(painted.diagnostics.some(a => a.reason === d.reason && a.path === d.path && a.pageId === d.pageId && a.severity === d.severity),
        `${name}: original input diagnostic retained at ${d.path}`);
    for (const d of paintOnly.diagnostics)
      assert.ok(painted.diagnostics.some(a => a.reason === d.reason && a.scenePath === d.scenePath && a.severity === d.severity));
    const outputDir = path.join(artifacts, `combined-${name}`);
    const publication = await publishPpjPreview(painted, previewInputEvidence(input, compiled), { outputDir });
    const persisted = JSON.parse(await readFile(path.join(outputDir, "render.json"), "utf8"));
    assert.deepEqual(persisted, JSON.parse(JSON.stringify(publication.receipt)));
    assert.deepEqual(persisted.assessment, JSON.parse(JSON.stringify(painted.assessment)));
    await assertPublishedWarning(persisted, painted);
    assert.equal(sha256(compiled.file), originalCandidate);
    assert.deepEqual(Buffer.from(compiled.programJson), originalInput);
    combinedAssessmentCases.push({ name, inputDiagnostics: painted.inputAssessment.diagnostics.length,
      reliability: painted.reliability.status, returnedPersistedEqual: true, warningPixels: true,
      candidateSha256: originalCandidate, sceneSha256: painted.sceneEvidence.sha256 });
  }
  for (const { name, input, compiled, painted } of publicationCases) {
    const evidence = previewInputEvidence(input, compiled);
    assert.throws(() => paintPpjSceneSvg({ ...compiled, previewScene: undefined, previewSceneBytes: undefined }),
      error => error.code === "preview.scene.missing");
    assert.throws(() => previewInputEvidence(input, { ...compiled, previewScene: { ...compiled.previewScene, version: 99 } }),
      error => error.code === "preview.scene.version");
    if (compiled.previewScene.assets.length) assert.throws(() => previewInputEvidence(input, { ...compiled, assets: [] }),
      error => error.code === "preview.scene.asset-mismatch");
    await assert.rejects(publishPpjPreview(painted, evidence, { outputDir: path.join(artifacts, `published-${name}`) }),
      error => error.code === "preview.output.exists");
    const beforeCandidate = sha256(compiled.file), beforeSource = input.source?.byteLength ? sha256(input.source) : null;
    const beforeAssessment = JSON.stringify(painted.assessment);
    for (const mode of ["load", "raster", "svg", "png", "pending", "final"]) {
      const outputDir = path.join(artifacts, `failure-${name}-${mode}`);
      let injected = 0;
      const options = mode === "load" ? { loadRaster: async () => {
        injected++; throw new Error("injected missing raster dependency");
      } } : mode === "raster" ? { loadRaster: async () => ({ render: async bytes => {
        if (!injected++) throw new Error("injected first-page raster failure");
        return sharp(bytes).png().toBuffer();
      } }) } : { writeArtifact: async (file, bytes) => {
        const base = path.basename(file);
        const matches = mode === "pending" ? base === "render.pending.json"
          : mode === "final" ? base === "render.json" : file.endsWith(`.${mode}`);
        if (matches && !injected++) throw Object.assign(new Error("injected disk full"), { code: "ENOSPC" });
        return writeExclusiveFile(file, bytes);
      } };
      let failure;
      await assert.rejects(publishPpjPreview(painted, evidence, { outputDir, ...options }), error => {
        failure = error;
        return error.code === (mode === "pending" ? "preview.output.pending"
          : mode === "final" ? "preview.output.manifest" : "preview.output.incomplete");
      });
      assert.ok(injected, "the requested failure must actually execute");
      const receipt = failure.receipt;
      assert.equal(receipt.ok, false);
      assert.equal(receipt.reliability.status, "failed");
      for (const [index, page] of receipt.pages.entries()) {
        assert.equal(page.reliability.status, ["load", "pending", "final"].includes(mode) || index === 0
          ? "failed" : painted.pages[index].reliability.status);
        for (const key of ["file", "png"]) if (page[key])
          assert.ok(receipt.artifacts.some(a => a.file === page[key] && a.pageId === page.id));
      }
      assert.deepEqual(receipt.scene, painted.sceneEvidence);
      assert.equal(receipt.compile.outputSha256, beforeCandidate);
      assert.equal(receipt.source?.sha256 ?? null, beforeSource);
      assert.deepEqual(receipt.assessment.children, receipt.pages.map(p => p.assessment));
      if (["pending", "final"].includes(mode)) {
        await assert.rejects(readFile(path.join(outputDir, "render.json")), { code: "ENOENT" });
        if (mode === "pending") await assert.rejects(readFile(path.join(outputDir, "render.pending.json")), { code: "ENOENT" });
        else {
          const pending = JSON.parse(await readFile(path.join(outputDir, "render.pending.json"), "utf8"));
          assert.equal(pending.ok, false);
          assert.equal(pending.output.status, "incomplete");
          assert.deepEqual(pending.scene, receipt.scene);
        }
      } else {
        assert.deepEqual(JSON.parse(await readFile(path.join(outputDir, "render.json"), "utf8")),
          JSON.parse(JSON.stringify(receipt)));
        await assert.rejects(readFile(path.join(outputDir, "render.pending.json")), { code: "ENOENT" });
      }
      for (const artifact of receipt.artifacts) {
        const bytes = await readFile(path.join(outputDir, artifact.file));
        assert.equal(sha256(bytes), artifact.sha256);
        if (artifact.kind === "png") assert.equal((await sharp(bytes).metadata()).format, "png");
      }
      assert.equal(sha256(compiled.file), beforeCandidate);
      assert.equal(input.source?.byteLength ? sha256(input.source) : null, beforeSource);
      assert.equal(JSON.stringify(painted.assessment), beforeAssessment);
      publicationFailures.push({ name, mode, code: failure.code, artifacts: receipt.artifacts.length,
        sceneSha256: receipt.scene.sha256, candidateSha256: beforeCandidate });
    }
  }
  const factualFailurePublications = [];
  for (const [name, input, compiled, expectedReason] of [
    ["scatter-connected", { ...sourceWorkspace, program: Buffer.from(JSON.stringify(scatterConnectedProgram)) },
      scatterConnected, "preview.scene.paint.scatter-line-unresolved"],
    ["diagram-import", { ...diagramInput, program: Buffer.from(JSON.stringify(diagramEdit)) },
      diagramCandidate, "preview.scene.paint.diagram-import-incomplete"],
  ]) {
    const painted = paintPpjSceneSvg(compiled);
    assert.ok(painted.diagnostics.some(d => d.reason === expectedReason));
    const outputDir = path.join(artifacts, `unavailable-${name}`);
    let receipt;
    await assert.rejects(publishPpjPreview(painted, previewInputEvidence(input, compiled), { outputDir }), error => {
      receipt = error.receipt;
      return error.code === "preview.output.incomplete";
    });
    assert.equal(receipt.ok, false);
    assert.equal(receipt.reliability.status, "failed");
    assert.deepEqual(receipt.scene, painted.sceneEvidence);
    assert.deepEqual(JSON.parse(await readFile(path.join(outputDir, "render.json"), "utf8")), JSON.parse(JSON.stringify(receipt)));
    await assertPublishedWarning(receipt, painted);
    for (const artifact of receipt.artifacts)
      assert.equal(sha256(await readFile(path.join(outputDir, artifact.file))), artifact.sha256);
    factualFailurePublications.push({ name, reason: expectedReason, redWarningPixels: true });
  }
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
  assert.deepEqual(await javascriptIdentity(), javascriptAtStart, "Evidence source files changed during integration; rerun a stable snapshot.");
  const nativeManifest = await readFile(path.join(path.dirname(packageJsonPath), "manifest.json"));
  assert.deepEqual(JSON.parse(nativeManifest.toString("utf8")), descriptors[0].manifest,
    "Package manifest changed during integration");
  for (const descriptor of descriptors) assert.equal(sha256(await readFile(descriptor.executablePath)),
    descriptor.manifest.files.find(file => file.path === descriptor.manifest.profiles[descriptor.profile].executable).sha256,
    "Executed package identity must still match its validated manifest");
  const report = { status: polygonPresetFailures.length || relationFailures.length || diagramFailures.length || transformProfileFailures.length || nestedPairFailures.length || datasetPairFailures.length || stylePairFailures.length || textAnchorFailures.length || spacingDeletionFailures.length || backgroundGradientFailures.length || backgroundImageFailures.length || shapeImageFailures.length || radialGradientFailures.length || brightnessFailures.length || textRotationFailures.length || textReflectionFailures.length || textDirectionFailures.length ? "failed" : "passed", scope: "PPJ NativeAOT wire/view, internal painting and selected production entry cases; not complete paint or installed-package acceptance",
    productionEntryCases,
    performance: { file: "performance.json", cases: performanceCases.map(c => ({ name: c.name, samples: c.samples.length })),
      sha256: sha256(await readFile(path.join(artifacts, "performance.json"))),
      scope: "warm size/time plus isolated native RSS/high-water; retained-object release and full 5.3/G-16 remain unverified" },
    stylePairFailures,
    datasetPairFailures,
    nestedPairFailures,
    transformProfileFailures,
    textAnchorFailures,
    textRotationCases,
    textRotationFailures,
    textRotationOpaqueCases,
    textReflectionCases,
    textReflectionFailures,
    textDirectionCases,
    textDirectionFailures,
    polygonPresetCases,
    polygonPresetFailures,
    spacingDeletionFailures,
    capitalizationCases,
    mediaPosterCases,
    gradientCases,
    radialGradientCases,
    radialGeometryCases,
    radialGradientFailures,
    brightnessCases,
    brightnessFailures,
    backgroundGradientCases,
    backgroundGradientFailures,
    backgroundImageCases,
    backgroundImageFailures,
    backgroundImageRejections,
    shapeImageCases,
    shapeImageFailures,
    characterBullets: { cases: characterBulletCases, reprojection: true, changedParts: bulletChangedParts },
    isolatedLineProfileCases,
    datasetProfileCases,
    spacingDeletions: { expectedCount: 7, cases: spacingDeletionCases },
    diagramFailures: diagramFailures.map(error => error.message),
    recordedAt: new Date().toISOString(),
    javascript: javascriptAtStart,
    runtimePackage: { manifestSha256: sha256(nativeManifest), packageVersion: descriptors[0].manifest.packageVersion,
      sdkVersion: descriptors[0].manifest.sdkVersion, target: descriptors[0].manifest.target },
    nativeBars, nativeStacks, nativeCircular,
    sceneAssessmentPublication: { authored: 2, sourceBound: 2, operationalFailures: publicationFailures,
      factualFailures: factualFailurePublications, nodeAssessmentBindings: true, diagnosticAddressesPreserved: true,
      visibilityProfile: { authored: true, sourceNoop: true, requiresMatchingPaint: true, otherRulesRetained: true },
      transformProfile: transformProfileCases,
      transformSourceEdits,
      combinedAssessment: combinedAssessmentCases,
      groupCoordinates: groupCoordinateCases,
      returnedPersistedEqual: true, artifactHashes: true, nonOverwrite: true, warningPixels: true,
      sceneIdentity: true, wrongCandidateRejected: true, corruptSceneAndAssetsRejected: true,
      scope: "shared scene identity/assessment/publisher and failure-path contract; production entry cases are recorded separately" },
    nativeScatter: { authored: true, authoredXChange: true, sourceNoop: true, sourceYEdit: true, sourceXEdit: "rejected as source-owned; explicit error asserted", numericPositionPixels: true, missingAndIsolatedPixels: true,
      reprojection: true, connectedMode: "unavailable: writer noFill verified; no invented SVG lines", changedParts: scatterChangedParts, sourceSha256: sha256(scatterSource), candidateSha256: sha256(scatterCandidate.file), workbook: "not present in literal-data fixture" },
    relationFailures: relationFailures.map(({ shift, actual, error }) => ({ shift, actual, message: error.message })),
    internalPainting: { artifacts, pairedComponent: 1,
      smallTextBaseline: { authoredAndSource: true, actualGlyphPixels: true, noDefaultSizeInflation: true, sourceNoop: true },
      textInsets: { authoredCases: 4, rightAlignedShift: -30, centeredShift: -15, identicalInkCount: true, fieldDiagnostics: true },
      textAnchors: { cases: textAnchorCases, edits: textAnchorEdits, expectedEditCount: 6,
        actualGlyphPixels: true, sourceNoop: true, approximateLineMetrics: true },
      styleGrammarPair: { fixture: "test/fixtures/presentation/preview-style-grammar-equivalence.json", status: stylePairFailures.length ? "failed" : "passed",
        ...(stylePairFailures.length ? {} : { visualPayloadBytesEqual: true, completeRasterEqual: true, sceneOnOffCandidateEqual: true,
          explicitFalseOverride: true, originalOwners: true, missingAndZero: true, resolvedColorPixels: true }) },
      nestedRepeatSlotPair: { fixture: "test/fixtures/presentation/preview-nested-repeat-equivalence.json", status: nestedPairFailures.length ? "failed" : "passed",
        variants: ["plain", "styled"], styleFixture: "test/fixtures/presentation/preview-style-grammar-equivalence.json",
        ...(nestedPairFailures.length ? {} : { visualPayloadBytesEqual: true, completeRasterEqual: true, sceneOnOffCandidateEqual: true,
          originalOwners: true, distinctInstances: true, missingAndZero: true }) },
      datasetPair: { fixture: "test/fixtures/presentation/preview-dataset-equivalence.json", status: datasetPairFailures.length ? "failed" : "passed",
        chartTypes: ["line", "heatmap"],
        ...(datasetPairFailures.length ? {} : { visualPayloadBytesEqual: true, completeRasterEqual: true, sceneOnOffCandidateEqual: true,
          multiSeries: true, originalOwners: true, missingAndZero: true }) },
      customArcPath, generatedBezierPaths: paths, sourceTextEdit: true, directedAnchorCases, requiredDirectedAnchorCases: 2,
      textFormats: { baselinePixels: true, decorationPixels: true, signedSpacingPixels: true, sourceNoop: true, sourceEditReprojection: true, changedParts: formatChangedParts },
      paragraphSpacing: { authored: true, sourceNoop: true, pixels: true, marginAndHangingPixels: true, zeroSpacingEdit: true, reprojection: true, changedParts: paragraphChangedParts, multiplierCases, paragraphMultiplierCases },
      shapeOutline: { authored: true, sourceNoop: true, dashPixels: true, sourceEditReprojection: true, changedParts: outlineChangedParts },
      diagramCache: { authoredPixels: true, sourceNoop: true, candidateText: true, reprojection: true, changedParts: diagramChangedParts,
        ownership: { source: diagramOldGraph.targets, candidate: diagramNewGraph.targets, outsideGraphPreserved: diagramFailures.length === 0 },
        importedPainting: "unavailable: importer omits cached connection and paint state; explicit failure asserted" },
      imageCrop: { positiveCrop: true, negativeLetterbox: true, sourceNoop: true, sourceEditReprojection: true, pixels: true },
      imageTile: { authoredAndSourceCases: 4, croppedAndUncropped: true, noInventedStretch: true,
        status: "unavailable: intrinsic tile sizing/DPI not resolved; not a tile rendering pass" },
      imageBorder: { authoredAndSourceEdit: true, rgbPixels: true, reprojection: true, nonTargetPreserved: true },
      imageMasks: { presets: ["ellipse", "diamond", "roundRect"], roundRectZeroEdit: true, customPath: true, authoredAndSourceEdit: true, cropCombinedPixels: true, reprojection: true },
      explicitCoordinateConnector: true, elbowBend: { authored: 25000, edited: 75000, sourceNoop: true, pixels: true, reprojection: true, changedParts: bendChangedParts },
      connectorProfile: { cases: connectorProfileCases, directEndpointsOnly: true, missingCaptureRetainsFailure: true },
      sourceConnectorPreserved: true, mergedTablePixels: true, sourceTableMoveReprojection: true },
    nativeLine: { authored: true, sourceNoop: true, preservedAcrossTableEdit: true, sourceValueEditReprojection: true,
      multiSeriesMissing: true, markerAndGapPixels: true, unclippedMarkerOutlinePixels: true, changedParts: chartChangedParts,
      sourceSha256: sha256(source), candidateSha256: sha256(chartCandidate.file), workbook: "not present in this literal-data fixture" },
    officeProfile: "presentation request rejected as designed", profiles: descriptors.map(d => ({
    profile: d.profile, path: d.executablePath,
    sha256: d.manifest.files.find(file => file.path === d.manifest.profiles[d.profile].executable).sha256,
  })), authored: 2, sourceBound: ["no-op with assets/table", "text leaf edit with assets/table", "table frame edit and fresh reprojection", "literal line value edit and fresh reprojection"] };
  await writeFile(path.join(artifacts, "integration.json"), JSON.stringify(report, null, 2), { flag: "wx" });
  console.log(JSON.stringify(report, null, 2));
  if (stylePairFailures.length) throw new AggregateError(stylePairFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Style/grammar equivalence failed; independent checks executed, not a passing integration.");
  if (datasetPairFailures.length) throw new AggregateError(datasetPairFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Dataset equivalence failed; independent checks executed, not a passing integration.");
  if (nestedPairFailures.length) throw new AggregateError(nestedPairFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Nested repeat equivalence failed; independent checks executed, not a passing integration.");
  if (backgroundGradientFailures.length) throw new AggregateError(backgroundGradientFailures.map(f => new Error(`${f.angle}: ${f.message}`)), "Background gradient regressions failed; independent checks executed, not a passing integration.");
  if (textRotationFailures.length) throw new AggregateError(textRotationFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Text rotation regressions failed; independent checks executed, not a passing integration.");
  assert.equal(textRotationCases.length, 47, "all text, shape and table authored/source angles, composed transforms and independent source edits must execute");
  if (textReflectionFailures.length) throw new AggregateError(textReflectionFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Text reflection regressions failed; independent checks executed, not a passing integration.");
  assert.equal(textReflectionCases.length, 36, "all shape/table self, group, nested and independently source-edited reflection cases must execute");
  if (textDirectionFailures.length) throw new AggregateError(textDirectionFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Text direction regressions failed; independent checks executed, not a passing integration.");
  assert.equal(textDirectionCases.length, 75, "all three text owners and directions, anchors, transforms and independent source edits must execute");
  if (polygonPresetFailures.length) throw new AggregateError(polygonPresetFailures.map(f => new Error(`${f.preset}/${f.consumer}/${f.stage}: ${f.message}`)), "Polygon preset regressions failed; independent checks executed.");
  assert.equal(polygonPresetCases.length, 66, "five presets, three consumers, authored/no-op and independent adjustment set/zero/delete edits");
  if (backgroundImageFailures.length) throw new AggregateError(backgroundImageFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Background image regressions failed; independent checks executed, not a passing integration.");
  if (shapeImageFailures.length) throw new AggregateError(shapeImageFailures.map(f => new Error(`${f.preset}/${f.edit ?? f.stage}: ${f.message}`)), "Shape image regressions failed; independent checks executed, not a passing integration.");
  if (radialGradientFailures.length) throw new AggregateError(radialGradientFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Radial gradient regressions failed; independent checks executed, not a passing integration.");
  assert.equal(radialGradientCases.length, 14, "all three consumers and their independent source edits must execute");
  assert.equal(radialGeometryCases.length, 8, "path bounds need native authored/source and independent source-edit evidence");
  if (brightnessFailures.length) throw new AggregateError(brightnessFailures.map(f => new Error(`${f.name}: ${f.message}`)), "Gradient brightness regressions failed; independent checks executed, not a passing integration.");
  assert.equal(brightnessCases.length, 36, "all consumers need both gradient kinds and independent source-stop edits");
  if (spacingDeletionFailures.length) throw new AggregateError(spacingDeletionFailures.map(f => new Error(`${f.mask}: ${f.message}`)), "Paragraph spacing deletion regressions failed; independent checks executed, not a passing integration.");
  assert.equal(spacingDeletionCases.length, 7);
  if (textAnchorFailures.length) throw new AggregateError(textAnchorFailures.map(f => new Error(`${f.operation}: ${f.message}`)), "Text anchor source regressions failed; independent checks executed, not a passing integration.");
  assert.equal(textAnchorEdits.length, 6, "all required text anchor edits must execute");
  if (transformProfileFailures.length) throw new AggregateError(transformProfileFailures.map(f => new Error(`${f.name}/${f.stage}: ${f.message}`)), "Transform source regressions failed; independent checks executed, not a passing integration.");
  if (diagramFailures.length) throw new AggregateError(diagramFailures, "Diagram source-edit part preservation needs ownership audit; independent regressions executed, not a passing integration.");
  if (relationFailures.length) throw new AggregateError(relationFailures.map(f => f.error), "Authored object-anchor regressions failed; independent table/source checks executed, not a passing integration.");
} finally {
  hook.deregister();
  delete globalThis[Symbol.for("officekit.preview.native.test")];
}
