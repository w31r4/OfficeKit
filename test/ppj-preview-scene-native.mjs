// Explicit integration against a checked-in-build-command output. This does not
// replace the installed package. Internal SVG foundations are exercised below;
// production scene routing/publication (tasks 3.3/4.x/5.2) remain separate.
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
  "../src/ppj/preview-scene.mjs", "../src/ppj/preview-scene-view.mjs", "../src/ppj/preview-diagnostics.mjs", "../src/ppj/preview-output.mjs",
  "../src/ppj/preview-input-assessment.mjs", "../src/ppj/preview-factual-errors.mjs", "../src/ppj/capability-registry.json",
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
  const xmlAttributes = tag => Object.fromEntries([...tag.matchAll(/([\w:]+)="([^"]*)"/g)].map(m => [m[1], m[2]]));
  // The controlled SDK fixture may reorder attributes and add a redundant
  // root declaration for the same a namespace already declared on children.
  // Do not remove arbitrary namespaces or alter any content/attribute values.
  const orderedXml = xml => xml.replace(/<p:sld\b[^>]*>/, tag => tag.replace(' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"', ""))
    .replace(/<([\w:.-]+)(\s[^<>]*?)?(\/?)>/g, (tag, name, _, close) =>
    `<${name} ${JSON.stringify(Object.entries(xmlAttributes(tag)).sort(([a],[b]) => a.localeCompare(b)))}${close}>`);
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
  for (const chartType of ["line", "heatmap"]) {
  const datasetPair = [];
  for (const name of ["encoded", "explicit"]) try {
    const program = structuredClone(pairBase);
    program.pages[0].elements = [{ id: "dataset-line", type: "chart", chartType, frame: datasetFixture.frame, data: datasetFixture[name] }];
    if (chartType === "heatmap") program.pages[0].elements[0].style = { heatmap: {
      colors: ["#000000", "#FF0000"], domain: [0, 5], missingFill: "#00FF00", showColorBar: false, showValues: false, cellGap: 0,
    } };
    const original = JSON.stringify(program);
    await writeFile(path.join(artifacts, `dataset-pair-${chartType}-${name}.ppj`), original, { flag: "wx" });
    const ordinary = await compilePpjWorkspace({ ...sourceWorkspace, program: Buffer.from(original) });
    const compiled = await compileFormat(program), painted = await savePaint(`dataset-pair-${chartType}-${name}`, compiled);
    assert.deepEqual(compiled.file, ordinary.file);
    const node = createPpjSceneView(compiled).pages[0].nodes[0];
    assert.deepEqual(node.frame, datasetFixture.frame);
    if (chartType === "line") {
    assert.equal(node.kind, "chart");
    assert.deepEqual(node.native.series.map(s => s.name), ["Alpha", "Beta"]);
    assert.deepEqual(node.native.series.map(s => s.values), [[1, 0, 0], [0, 4, 5]]);
    assert.deepEqual(node.native.series.map(s => s.missingValueIndexes), [[1], []]);
    assert.equal((painted.pages[0].svg.match(/data-officekit-review-point="isolated"/g) || []).length, 2);
    assert.equal((painted.pages[0].svg.match(/data-officekit-line-segment=/g) || []).length, 1);
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
  } catch (error) {
    datasetPairFailures.push({ chartType, name, code: error.code, message: error.message });
    console.error(`Dataset pair ${name} failed: ${error.message}`);
  }
  if (datasetPair.length === 2) try {
    assert.deepEqual(datasetPair[0].visual, datasetPair[1].visual);
    assert.deepEqual(datasetPair[0].raster, datasetPair[1].raster);
  } catch (error) { datasetPairFailures.push({ chartType, name: "equivalence", code: error.code, message: error.message }); }
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
    for (const diagnostic of canonical.diagnostics.filter(d => d.reason !== reason))
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
  const report = { status: relationFailures.length || diagramFailures.length || transformProfileFailures.length || nestedPairFailures.length || datasetPairFailures.length || stylePairFailures.length || textAnchorFailures.length || spacingDeletionFailures.length ? "failed" : "passed", scope: "PPJ NativeAOT wire/view and internal SVG foundations; not production scene routing or complete paint coverage",
    performance: { file: "performance.json", cases: performanceCases.map(c => ({ name: c.name, samples: c.samples.length })),
      sha256: sha256(await readFile(path.join(artifacts, "performance.json"))),
      scope: "warm size/time plus isolated native RSS/high-water; retained-object release and full 5.3/G-16 remain unverified" },
    stylePairFailures,
    datasetPairFailures,
    nestedPairFailures,
    transformProfileFailures,
    textAnchorFailures,
    spacingDeletionFailures,
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
      scope: "internal scene identity/assessment/publisher and failure-path contract; production routing remains open" },
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
