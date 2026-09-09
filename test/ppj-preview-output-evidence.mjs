import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, readdir, rm, symlink, writeFile } from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import { registerHooks } from "node:module";
import { publishPpjPreview, previewInputEvidence, previewPageStems, PreviewOutputError } from "../src/ppj/preview-output.mjs";
import { renderPpjToSvg } from "../src/ppj/svg-preview.mjs";
import { sha256, writeExclusiveFile } from "../src/ppj/workspace.mjs";
import { previewAssessment } from "../src/ppj/preview-diagnostics.mjs";
import { ppjPreviewSceneIdentity } from "../src/ppj/preview-scene.mjs";

const root = await mkdtemp(path.join(os.tmpdir(), "officekit-preview-evidence-"));
const page = (id) => ({ id, svg: `<svg xmlns="http://www.w3.org/2000/svg"><text>${id}</text></svg>`, diagnostics: [] });
const result = {
  renderer: "officekit-svg-preview", canvas: { width: 100, height: 100 },
  pages: [page("first"), page("second")], diagnostics: [], status: "supported",
};
const workspace = {
  path: path.join(root, "input.ppj"), program: Buffer.from("program"), assets: [],
  source: Buffer.from("source PPTX bytes"), sourcePath: path.join(root, "source.pptx"),
};
const compiled = { programSha256: sha256("canonical"), file: Buffer.from("different candidate bytes"), sourceBound: true };
const evidence = previewInputEvidence(workspace, compiled);
// These bytes test publication only, not image correctness. Real sharp/codec
// output is checked separately by ppj-svg-preview.mjs.
const loadRaster = async () => ({ identity: { name: "test-raster", status: "available" }, render: async (svg) => Buffer.concat([Buffer.from("test-png:"), svg]) });
const publish = (name, options = {}, value = result) => publishPpjPreview(value, evidence, { outputDir: path.join(root, name), loadRaster, ...options });
const json = async (name) => JSON.parse(await readFile(path.join(root, name, "render.json"), "utf8"));
async function rejected(promise, code) {
  let caught;
  await assert.rejects(promise, (error) => {
    caught = error;
    assert.equal(error.receipt.reliability.status, "failed", "every production failure must fail reliability");
    assert.equal(error.receipt.assessment.status, "unavailable");
    assert.deepEqual(error.receipt.assessment.reliability, error.receipt.reliability);
    return error instanceof PreviewOutputError && error.code === code;
  });
  return caught;
}
async function verifyArtifacts(receipt) {
  for (const artifact of receipt.artifacts) {
    const bytes = await readFile(path.join(receipt.output.directory, artifact.file));
    assert.equal(bytes.byteLength, artifact.bytes);
    assert.equal(sha256(bytes), artifact.sha256);
  }
  for (const p of receipt.pages) {
    for (const [key, kind] of [["file", "svg"], ["png", "png"]]) {
      if (p[key]) assert.ok(receipt.artifacts.some((a) => a.file === p[key] && a.kind === kind && a.pageId === p.id));
    }
  }
}

try {
  await writeFile(workspace.path, workspace.program);
  await writeFile(workspace.sourcePath, workspace.source);
  const assetPath = path.join(root, "asset.png");
  await writeFile(assetPath, "original asset");
  await mkdir(path.join(root, "existing"));
  await writeFile(path.join(root, "existing", "first.svg"), "keep this evidence");
  for (const target of ["existing", "input.ppj", "source.pptx", "asset.png"]) {
    await rejected(publish(target), "preview.output.exists");
  }
  assert.equal(await readFile(path.join(root, "existing", "first.svg"), "utf8"), "keep this evidence");
  assert.deepEqual(await readFile(workspace.path), workspace.program);
  assert.deepEqual(await readFile(workspace.sourcePath), workspace.source);
  assert.equal(await readFile(assetPath, "utf8"), "original asset");
  await symlink(path.join(root, "existing"), path.join(root, "linked"), "dir");
  await rejected(publish("linked"), "preview.output.exists");
  await symlink(path.join(root, "absent-target"), path.join(root, "dangling"), "dir");
  await rejected(publish("dangling"), "preview.output.exists");
  await assert.rejects(readFile(path.join(root, "absent-target", "render.json")), { code: "ENOENT" });

  const race = await Promise.allSettled([publish("race"), publish("race")]);
  assert.equal(race.filter((r) => r.status === "fulfilled").length, 1);
  assert.equal(race.find((r) => r.status === "rejected").reason.code, "preview.output.exists");
  await verifyArtifacts(await json("race"));

  const badIds = ["../escape", "..\\escape", "/absolute", "CON", "A", "a", "slide-1", "", "nul", "LPT1", "a.b"];
  const stems = previewPageStems(badIds.map(page));
  assert.equal(new Set(stems.map((s) => s.toLowerCase())).size, stems.length);
  assert.ok(stems.every((s) => /^[a-z0-9_-]+$/iu.test(s)));
  const { receipt: mapped } = await publish("mapped", {}, { ...result, pages: badIds.map(page) });
  assert.deepEqual(mapped.pages.map((p) => p.id), badIds);
  await verifyArtifacts(mapped);
  assert.equal((await readdir(path.join(root, "mapped"))).length, badIds.length * 2 + 1);
  await assert.rejects(readFile(path.join(root, "escape.svg")), { code: "ENOENT" });

  // Successful real filesystem publication and input/candidate distinctions.
  const { receipt } = await publish("success");
  assert.deepEqual(receipt, await json("success"));
  assert.equal(receipt.ok, true);
  assert.equal(receipt.output.status, "complete");
  assert.equal(receipt.artifacts.length, 4);
  assert.equal(receipt.sourceBound, true);
  assert.equal(receipt.source.sha256, sha256(workspace.source));
  assert.equal(receipt.compile.outputSha256, sha256(compiled.file));
  assert.notEqual(receipt.source.sha256, receipt.compile.outputSha256);
  assert.equal(receipt.input.sha256, sha256(workspace.program));
  assert.equal(receipt.visualReview, "requires-human");
  assert.equal(receipt.reliability.status, "requires-review", "unassessed publisher input is not a visual pass");
  assert.equal(receipt.environment.node, process.version);
  for (const source of receipt.implementation.sources) {
    assert.equal(source.sha256, sha256(await readFile(new URL(`../src/ppj/${source.file}`, import.meta.url))));
  }
  await verifyArtifacts(receipt);
  await assert.rejects(readFile(path.join(root, "success", "render.pending.json")), { code: "ENOENT" });

  const missing = await rejected(publish("missing", { loadRaster: async () => { throw new Error("test dependency missing"); } }), "preview.output.incomplete");
  assert.equal(missing.receipt.ok, false);
  assert.equal(missing.receipt.status, "unavailable");
  assert.equal(missing.receipt.environment.raster.status, "unavailable");
  assert.equal(missing.receipt.failures[0].stage, "raster-load");
  assert.ok(missing.receipt.pages.every((p) => p.file && !("png" in p)));
  assert.deepEqual(missing.receipt, await json("missing"));
  await verifyArtifacts(missing.receipt);

  for (const failedIndex of [0, 1]) {
    let index = 0;
    const error = await rejected(publish(`raster-${failedIndex}`, { loadRaster: async () => ({ render: async () => {
      if (index++ === failedIndex) throw new Error("raster execution failed");
      return Buffer.from("test PNG");
    } }) }), "preview.output.incomplete");
    assert.equal(error.receipt.failures[0].stage, "raster-render");
    assert.equal(error.receipt.failures[0].pageId, result.pages[failedIndex].id);
    assert.equal(error.receipt.artifacts.length, 3);
    assert.ok(!error.receipt.pages[failedIndex].png);
    await verifyArtifacts(error.receipt);
  }
  const emptyRaster = await rejected(publish("empty-raster", { loadRaster: async () => ({ render: async () => Buffer.alloc(0) }) }), "preview.output.incomplete");
  assert.ok(emptyRaster.receipt.failures.every((f) => f.stage === "raster-render"));

  // Lowered pages share the semantic root but have distinct native addresses.
  // A failed page must not replace its siblings in the document assessment.
  const scenePages = result.pages.map((p, index) => {
    const assessment = previewAssessment({ path: "$", pageId: p.id,
      scenePath: `$.presentation.slides[${index}]`, assessed: true });
    return { ...p, assessment, diagnostics: assessment.diagnostics };
  });
  const sceneResult = { ...result, pages: scenePages,
    assessment: previewAssessment({ path: "$", scenePath: "$.presentation", assessed: true,
      children: scenePages.map(p => p.assessment) }) };
  for (const failedIndex of [0, 1]) {
    let index = 0;
    const name = `scene-page-failure-${failedIndex}`;
    const error = await rejected(publish(name, { loadRaster: async () => ({ render: async () => {
      if (index++ === failedIndex) throw new Error("page-specific raster failure");
      return Buffer.from("test PNG");
    } }) }, sceneResult), "preview.output.incomplete");
    const receipt = error.receipt;
    assert.deepEqual(receipt.assessment.children, receipt.pages.map(p => p.assessment));
    assert.deepEqual(receipt.assessment.children.map(p => p.pageId), ["first", "second"]);
    for (const [i, p] of receipt.pages.entries()) {
      assert.equal(p.reliability.status, i === failedIndex ? "failed" : "passed");
      for (const d of p.diagnostics) {
        assert.equal(d.pageId, p.id);
        assert.equal(d.scenePath, p.assessment.scenePath);
      }
    }
    assert.deepEqual(receipt, await json(name));
    await verifyArtifacts(receipt);
  }
  assert.ok(sceneResult.assessment.children.every(p => p.reliability.status === "passed"),
    "publication failure must not mutate the input assessment");
  for (const mode of ["global", "both-pages", "write"]) {
    const name = `scene-failure-${mode}`;
    const options = mode === "write" ? { writeArtifact: async (file, bytes) => {
      if (file.endsWith("first.png")) throw Object.assign(new Error("disk full"), { code: "ENOSPC" });
      return writeExclusiveFile(file, bytes);
    } } : { loadRaster: async () => {
      if (mode === "global") throw new Error("raster backend missing");
      return { render: async () => { throw new Error("raster execution failed"); } };
    } };
    const { receipt } = await rejected(publish(name, options, sceneResult), "preview.output.incomplete");
    assert.deepEqual(receipt.assessment.children, receipt.pages.map(p => p.assessment));
    assert.deepEqual(receipt, await json(name));
    assert.deepEqual(receipt.pages.map(p => p.reliability.status),
      mode === "write" ? ["failed", "passed"] : ["failed", "failed"]);
    const failures = receipt.diagnostics.filter(d => d.reason.startsWith("preview.output.") || d.reason === "raster-dependency-unavailable");
    assert.deepEqual(failures.map(d => d.scenePath).sort(), mode === "global" ? ["$.presentation"]
      : mode === "write" ? ["$.presentation.slides[0]"] : ["$.presentation.slides[0]", "$.presentation.slides[1]"]);
    await verifyArtifacts(receipt);
  }

  // Exclusive write collision during publication must preserve the other file.
  const collision = await rejected(publish("collision", { writeArtifact: async (file, bytes) => {
    if (file.endsWith("first.svg")) await writeFile(file, "concurrent external file", { flag: "wx" });
    return writeExclusiveFile(file, bytes);
  } }), "preview.output.incomplete");
  assert.equal(await readFile(path.join(root, "collision", "first.svg"), "utf8"), "concurrent external file");
  assert.equal(collision.receipt.failures[0].stage, "artifact-write");
  assert.ok(!collision.receipt.pages[0].file);

  for (const [name, failFile, code] of [
    ["pending-fail", "render.pending.json", "preview.output.pending"],
    ["manifest-fail", "render.json", "preview.output.manifest"],
    ["svg-fail", "first.svg", "preview.output.incomplete"],
    ["png-fail", "first.png", "preview.output.incomplete"],
  ]) {
    const error = await rejected(publish(name, { writeArtifact: async (file, bytes) => {
      if (path.basename(file) === failFile) throw Object.assign(new Error("injected disk failure"), { code: "ENOSPC" });
      return writeExclusiveFile(file, bytes);
    } }), code);
    assert.equal(error.receipt.output.status, "incomplete");
    assert.equal(error.receipt.ok, false);
    if (name === "pending-fail") assert.deepEqual(await readdir(path.join(root, name)), []);
    else if (name === "manifest-fail") {
      assert.equal(JSON.parse(await readFile(path.join(root, name, "render.pending.json"))).output.status, "incomplete");
      await assert.rejects(readFile(path.join(root, name, "render.json")), { code: "ENOENT" });
    } else {
      assert.ok(!error.receipt.artifacts.some((a) => a.file === failFile));
      await verifyArtifacts(error.receipt);
    }
  }
  const cleanup = await publish("cleanup", { removePending: async () => { throw new Error("cleanup denied"); } });
  assert.equal(cleanup.warnings[0].stage, "pending-cleanup");
  assert.deepEqual(cleanup.receipt, await json("cleanup"));
  assert.equal(cleanup.receipt.ok, true);
  assert.ok(await readFile(path.join(root, "cleanup", "render.pending.json")));

  // These are identity-contract fixtures, not validated native paint fixtures.
  // The real runtime test independently validates the scene digest and bytes.
  const scene = { version: 1, origin: 2, sha256: sha256("scene"),
    programSha256: compiled.programSha256, candidateSha256: sha256(compiled.file) };
  const boundResult = { ...result, renderer: "officekit-native-scene-svg-internal", scene,
    sceneEvidence: ppjPreviewSceneIdentity(scene) };
  const boundEvidence = { ...evidence, scene: boundResult.sceneEvidence };
  for (const key of ["version", "origin", "sha256", "programSha256", "candidateSha256", "missing"]) {
    const name = `identity-${key}`;
    const changed = key === "missing" ? undefined : { ...boundEvidence.scene,
      [key]: key === "version" ? 99 : key === "origin" ? "authored-lowering" : sha256("wrong") };
    const error = await rejected(publishPpjPreview(boundResult, { ...boundEvidence, scene: changed }, {
      outputDir: path.join(root, name), loadRaster: async () => { assert.fail("no raster before identity verification"); },
    }), "preview.output.scene");
    assert.equal(error.receipt.failures[0].stage, "scene");
    assert.deepEqual(error.receipt.artifacts, []);
    await assert.rejects(readdir(path.join(root, name)), { code: "ENOENT" });
  }
  const bound = await publishPpjPreview(boundResult, boundEvidence, { outputDir: path.join(root, "identity-success"), loadRaster });
  assert.deepEqual(bound.receipt.scene, boundResult.sceneEvidence);
  assert.deepEqual(bound.receipt, await json("identity-success"));
  for (const [name, value, inputEvidence] of [
    ["missing-painted", { ...boundResult, scene: undefined }, boundEvidence],
    ["missing-capture", { ...boundResult, sceneEvidence: undefined }, boundEvidence],
    ["wrong-compile", boundResult, { ...boundEvidence, compile: { ...evidence.compile, outputSha256: sha256("wrong") } }],
    ["wrong-mode", boundResult, { ...boundEvidence, sourceBound: false }],
    ["invalid-version", { ...boundResult, scene: { ...scene, version: 99 } }, boundEvidence],
  ]) {
    await rejected(publishPpjPreview(value, inputEvidence, {
      outputDir: path.join(root, name), loadRaster: async () => { assert.fail("invalid scene must not load raster"); },
    }), "preview.output.scene");
    await assert.rejects(readdir(path.join(root, name)), { code: "ENOENT" });
  }

  // Exercise the preview boundary without a native runtime: compilation sees
  // loaded bytes, then the path is changed before the SVG is built.
  const originalAsset = Buffer.from("asset snapshot");
  await writeFile(assetPath, originalAsset);
  const imageProgram = Buffer.from(JSON.stringify({
    design: { canvas: { width: 20, height: 20 } },
    pages: [{ id: "image-page", elements: [{ id: "image", type: "image", asset: "asset", frame: { x: 0, y: 0, width: 20, height: 20 } }] }],
  }));
  const loaded = { ...workspace, program: imageProgram, assets: [{ id: "asset", data: originalAsset, mimeType: "image/png" }] };
  let loads = 0;
  const snapshot = await renderPpjToSvg("unused.ppj", {
    load: async () => loaded,
    compile: async (w) => {
      assert.deepEqual(w.assets[0].data, originalAsset);
      await writeFile(assetPath, "changed later");
      return { ...compiled, programJson: imageProgram };
    },
    loadRaster: async () => { loads++; throw new Error("must remain lazy"); },
  });
  assert.equal(loads, 0);
  assert.match(snapshot.pages[0].svg, new RegExp(originalAsset.toString("base64")));
  assert.equal(previewInputEvidence(loaded, compiled).assets[0].sha256, sha256(originalAsset));
  for (const data of [[], [{ id: "asset", data: Buffer.alloc(0), mimeType: "image/png" }]]) {
    const unavailable = await renderPpjToSvg("unused.ppj", {
      load: async () => ({ ...loaded, assets: data }), compile: async () => ({ ...compiled, programJson: imageProgram }),
    });
    assert.ok(unavailable.diagnostics.some((d) => d.reason === "asset-missing"
      && d.path === "$.pages[0].elements[0].asset" && d.severity === "error" && d.status === "unavailable"));
    assert.equal(unavailable.reliability.status, "failed");
    assert.doesNotMatch(unavailable.pages[0].svg, /href=""/);
    const error = await rejected(publish(`asset-${data.length}`, {}, unavailable), "preview.output.incomplete");
    assert.equal(error.receipt.failures[0].stage, "asset");
  }

  // Presence edits must not silently claim that the local SVG draws error
  // bars. Codec/projection behavior is covered by PpjErrorBarsLifecycle.
  for (const chartType of ["column", "bar", "line", "combo"]) {
    for (const errorBars of [undefined, { valueType: "fixed-value", value: 0 },
      { valueType: "custom", plus: { values: [0, 1] }, minus: { values: [1, 0] } }, undefined]) {
      const programJson = Buffer.from(JSON.stringify({ pages: [{ id: "chart-page", elements: [{
        id: "errors", type: "chart", chartType, frame: { x: 0, y: 0, width: 200, height: 150 },
        data: { categories: ["A", "B"], series: [{ chartType: "line", values: [1, 2], errorBars }] },
      }] }] }));
      const preview = await renderPpjToSvg("unused.ppj", {
        load: async () => ({ ...workspace, program: programJson }),
        compile: async () => ({ ...compiled, programJson }),
      });
      const diagnostic = preview.diagnostics.find((d) => d.id === "errors" && d.reason === "chart-error-bars-not-rendered");
      if (errorBars) {
        assert.equal(diagnostic.status, "partial");
        assert.equal(diagnostic.reason, "chart-error-bars-not-rendered");
      } else assert.equal(diagnostic, undefined);
    }
  }

  // A root import must not resolve a specialist rendering/native dependency.
  const hook = registerHooks({ resolve(specifier, context, nextResolve) {
    assert.doesNotMatch(specifier, /^(sharp|mupdf|@office-kit\/codec-)/);
    return nextResolve(specifier, context);
  } });
  try { await import("../src/index.mjs"); } finally { hook.deregister(); }
  console.log("ppj preview output evidence ok: preservation, races, paths, partial writes, hashes, assets and lazy loading");
} finally {
  await rm(root, { recursive: true, force: true });
}
