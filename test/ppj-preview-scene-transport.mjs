import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { registerHooks } from "node:module";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import { BinaryWriter, WireType } from "@bufbuild/protobuf/wire";
import {
  CodecRequestSchema, CodecResponseSchema, PresentationPreviewSceneSchema,
  PresentationProgramResultSchema,
  PresentationElementSourceBindingSchema,
} from "../src/generated/office_kit/artifact/v1/office_artifact_pb.js";
import { encodePpjCodecRequest, decodePpjCodecResponse } from "../src/codecs/office-kit-ppj-wire.mjs";
import { readPpjPreviewScene } from "../src/ppj/preview-scene.mjs";

const hash = (bytes) => createHash("sha256").update(bytes).digest("hex");
const text = (value) => new TextEncoder().encode(value);
function seal(scene) { scene.sha256 = hash(toBinary(PresentationPreviewSceneSchema, { ...scene, sha256: "" })); }
function fixture(sourceBound = false) {
  // Synthetic transport bytes; these are not claimed to be a rendered PPTX.
  const file = text("candidate bytes for transport assertions");
  const programJson = text('{"pages":[{"id":"page","elements":[]}]}');
  const asset = text("asset bytes for identity assertions, not a decoded image");
  const program = create(PresentationProgramResultSchema, {
    programJson, programSha256: hash(programJson), outputSha256: hash(file), sourceBound,
    assets: [{ id: "ppj-source-image", contentType: "image/png", sha256: hash(asset), data: asset }],
    previewScene: {
      version: 1, origin: sourceBound ? 2 : 1, programSha256: hash(programJson), candidateSha256: hash(file),
      presentation: { id: "deck", slideWidthEmu: 12192000n, slideHeightEmu: 6858000n, slides: [{ id: "native-page", elements: [
        { id: "native-chart", hidden: false, content: { case: "chart", value: {
          yAxis: { minimum: 0, reverse: false }, series: [{ name: "A", values: [1, 0, 0], missingValueIndexes: [1] }],
        } } },
        { id: "native-image", content: { case: "image", value: { assetId: "native-asset" } } },
      ] }] },
      bindings: ["chart", "image"].map((name, i) => ({ pageId: "page", semanticId: name, nativeId: `native-${name}`,
        programPath: `$.pages[0].elements[${i}]`, scenePath: `$.presentation.slides[0].elements[${i}]`, attribution: 1, zOrder: i })),
      assets: [{ nativeId: "native-asset", contentType: "image/png", sha256: hash(asset) }],
    },
  });
  seal(program.previewScene);
  return { file, program };
}

function rejects(change, code, { reseal = true } = {}) {
  const value = fixture();
  change(value);
  if (reseal && value.program.previewScene) seal(value.program.previewScene);
  assert.throws(() => readPpjPreviewScene(value.program, value.file), (error) => error.code === `preview.scene.${code}`);
}

if (process.argv.includes("--native-forwarding")) {
  let next;
  const calls = [];
  globalThis[Symbol.for("officekit.preview.transport.test")] = async (factory, options) => {
    const request = factory();
    calls.push(fromBinary(CodecRequestSchema, encodePpjCodecRequest(request)));
    assert.equal(options.fileSidecar, true);
    const response = next();
    // Exercise the actual lightweight result decoder before native.mjs consumes
    // it, rather than handing a pre-decoded scene directly to the facade.
    return options.consumeResponse(decodePpjCodecResponse(toBinary(CodecResponseSchema, response)));
  };
  const stub = `export const OFFICE_KIT_PROTOCOL_VERSION=2; export const codecLimits=(x)=>x;
    export const invokeOfficeKitPpjLazy=(...args)=>globalThis[Symbol.for("officekit.preview.transport.test")](...args);`;
  const hook = registerHooks({ resolve(specifier, context, resolve) {
    if (specifier === "../codecs/office-kit-runtime.mjs" && context.parentURL.endsWith("/src/ppj/native.mjs"))
      return { url: `data:text/javascript,${encodeURIComponent(stub)}`, shortCircuit: true };
    return resolve(specifier, context);
  } });
  try {
    const { compilePpjWorkspace, validatePpjWorkspace } = await import("../src/ppj/workspace.mjs");
    const { compilePpjToPptx } = await import("../src/ppj/native.mjs");
    const workspace = { program: text("input program"), source: new Uint8Array(), assets: [] };
    const response = (include = true) => {
      const { file, program } = fixture();
      if (!include) program.previewScene = undefined;
      return create(CodecResponseSchema, { protocolVersion: 2, ok: true, file, presentationProgram: program });
    };
    next = () => response(false);
    assert.equal(Object.hasOwn(await compilePpjWorkspace(workspace), "previewScene"), false);
    assert.equal(calls.at(-1).presentationProgram.includePreviewScene, false);
    assert.equal(Object.hasOwn(await validatePpjWorkspace(workspace), "previewScene"), false);
    assert.equal(calls.at(-1).presentationProgram.validationOnly, true);
    assert.equal(calls.at(-1).presentationProgram.includePreviewScene, false);
    next = () => response();
    const receipt = await compilePpjWorkspace(workspace, { includePreviewScene: true, includeNodeMap: false });
    assert.equal(calls.at(-1).presentationProgram.includePreviewScene, true);
    assert.equal(calls.at(-1).presentationProgram.includeNodeMap, false);
    assert.equal(receipt.previewScene.presentation.slides[0].elements[0].content.value.yAxis.minimum, 0);
    assert.equal(receipt.assets[0].id, "ppj-source-image");
    assert.equal(receipt.previewScene.assets[0].nativeId, "native-asset");
    next = () => response(false);
    await assert.rejects(compilePpjWorkspace(workspace, { includePreviewScene: true }), { code: "preview.scene.missing" });
    next = () => { const r = response(); r.file = text("changed candidate"); return r; };
    await assert.rejects(compilePpjWorkspace(workspace, { includePreviewScene: true }), { code: "preview.scene.candidate-mismatch" });
    const count = calls.length;
    await assert.rejects(compilePpjToPptx(workspace.program, { includePreviewScene: true, validationOnly: true }), /requires compilation/u);
    assert.equal(calls.length, count, "invalid operation must not invoke the runtime");
  } finally {
    hook.deregister();
    delete globalThis[Symbol.for("officekit.preview.transport.test")];
  }
} else {
  for (const includePreviewScene of [false, true]) {
    const request = { protocolVersion: 2, operation: 11, family: 2,
      presentationProgram: { programJson: text("input"), includePreviewScene } };
    assert.deepEqual(encodePpjCodecRequest(request), toBinary(CodecRequestSchema, create(CodecRequestSchema, request)));
    assert.equal(fromBinary(CodecRequestSchema, encodePpjCodecRequest(request)).presentationProgram.includePreviewScene, includePreviewScene);
  }
  for (const sourceBound of [false, true]) {
    const { program, file } = fixture(sourceBound);
    const bytes = toBinary(CodecResponseSchema, create(CodecResponseSchema, { ok: true, file, presentationProgram: program }));
    const general = fromBinary(CodecResponseSchema, bytes);
    const ppj = decodePpjCodecResponse(bytes);
    assert.ok(ppj.presentationProgram.previewSceneBytes instanceof Uint8Array);
    const before = toBinary(PresentationProgramResultSchema, general.presentationProgram);
    const scene = readPpjPreviewScene(general.presentationProgram, general.file);
    assert.deepEqual(readPpjPreviewScene(ppj.presentationProgram, ppj.file), scene);
    assert.deepEqual(toBinary(PresentationProgramResultSchema, general.presentationProgram), before);
    assert.ok(Object.isFrozen(scene.presentation.slides[0].elements));
    const chart = scene.presentation.slides[0].elements[0].content.value;
    assert.equal(chart.yAxis.minimum, 0);
    assert.equal(chart.yAxis.reverse, false);
    assert.equal(chart.yAxis.maximum, undefined);
    assert.deepEqual(chart.series[0].values, [1, 0, 0]);
    assert.deepEqual(chart.series[0].missingValueIndexes, [1]);
    assert.equal(scene.presentation.slideWidthEmu, 12192000n);
  }
  rejects(v => { v.program.previewScene = undefined; }, "missing");
  rejects(v => { v.program.previewScene.version = 2; }, "version");
  rejects(v => { v.program.previewScene.origin = 0; }, "origin");
  rejects(v => { v.program.previewScene.origin = 2; }, "origin");
  rejects(v => { v.program.programJson = text("different canonical program"); }, "program-mismatch");
  rejects(v => { v.program.programSha256 = "a".repeat(64); }, "program-mismatch");
  rejects(v => { v.program.previewScene.candidateSha256 = "a".repeat(64); }, "candidate-mismatch");
  rejects(v => { v.file = text("different candidate"); }, "candidate-mismatch");
  rejects(v => { v.program.previewScene.presentation.slides[0].elements[0].hidden = true; }, "digest-mismatch", { reseal: false });
  rejects(v => { v.program.previewScene.bindings.pop(); }, "bindings");
  rejects(v => { v.program.previewScene.bindings.push(v.program.previewScene.bindings[0]); }, "bindings");
  rejects(v => { v.program.previewScene.bindings[0].nativeId = "other"; }, "bindings");
  rejects(v => { v.program.previewScene.bindings[0].zOrder = 9; }, "bindings");
  rejects(v => { v.program.previewScene.bindings[0].programPath = "$.pages[0]#component(fake)"; }, "bindings");
  rejects(v => { v.program.assets[0].data = text("corrupt asset bytes"); }, "asset-mismatch");
  rejects(v => { v.program.assets[0].data = new Uint8Array(); }, "asset-mismatch");
  rejects(v => { v.program.previewScene.assets[0].contentType = "image/jpeg"; }, "asset-mismatch");
  rejects(v => { v.program.previewScene.assets.push(v.program.previewScene.assets[0]); }, "asset-mismatch");
  rejects(v => { v.program.previewScene.assets = []; }, "asset-mismatch");
  rejects(v => { v.program.previewScene.presentation.slides[0].elements[1].content.value.assetId = "missing"; }, "asset-mismatch");
  rejects(v => { v.program.previewScene.presentation.slides[0].elements[0].source = create(PresentationElementSourceBindingSchema); }, "authority");
  const value = fixture();
  value.program.assets[0].sha256 = value.program.assets[0].sha256.toUpperCase();
  value.program.assets[0].contentType = "IMAGE/PNG";
  assert.ok(readPpjPreviewScene(value.program, value.file), "native MIME/hash matching is case insensitive");
  assert.throws(() => readPpjPreviewScene(value.program, value.file, { limits: { maxUncompressedBytes: 1 } }), { code: "preview.scene.budget" });
  assert.throws(() => readPpjPreviewScene({ previewSceneBytes: Uint8Array.of(0xff) }, value.file), { code: "preview.scene.invalid" });
  const invalidField = new BinaryWriter().tag(13, WireType.Varint).uint32(1).finish();
  const invalidResponse = new BinaryWriter().tag(9, WireType.LengthDelimited).bytes(invalidField).finish();
  assert.throws(() => decodePpjCodecResponse(invalidResponse), /Unexpected protobuf wire type/u);

  const child = spawnSync(process.execPath, [fileURLToPath(import.meta.url), "--native-forwarding"], { encoding: "utf8" });
  assert.equal(child.status, 0, child.stderr || child.stdout);
  // A clean process proves ordinary native/workspace imports and calls do not
  // load the full schema, preview decoder or raster backend.
  const lazy = spawnSync(process.execPath, ["--input-type=module", "-e", `
    import assert from 'node:assert/strict'; import { registerHooks } from 'node:module';
    const stub = 'export const OFFICE_KIT_PROTOCOL_VERSION=2; export const codecLimits=x=>x; export async function invokeOfficeKitPpjLazy(factory, options) { const r=factory(); if(r.presentationProgram.includePreviewScene) throw Error("scene unexpectedly requested"); return options.consumeResponse({file:new Uint8Array(),diagnostics:[],presentationProgram:{changedParts:[],changedNodeIds:[],assets:[]}}); }';
    const hook=registerHooks({resolve(specifier,context,next){
      assert.doesNotMatch(specifier,/preview-scene|generated\\/office_kit|^(sharp|mupdf)$/);
      if(specifier==='../codecs/office-kit-runtime.mjs' && context.parentURL.endsWith('/src/ppj/native.mjs')) return {url:'data:text/javascript,'+encodeURIComponent(stub),shortCircuit:true};
      return next(specifier,context);
    }});
    const {compilePpjWorkspace,validatePpjWorkspace}=await import(${JSON.stringify(new URL("../src/ppj/workspace.mjs", import.meta.url).href)});
    const w={program:new Uint8Array([1]),source:new Uint8Array(),assets:[]};
    await compilePpjWorkspace(w); await validatePpjWorkspace(w); hook.deregister();
  `], { encoding: "utf8" });
  assert.equal(lazy.status, 0, lazy.stderr || lazy.stdout);
  console.log("ppj preview scene transport ok: both wire profiles, opt-in forwarding, identities, assets, presence and lazy defaults");
}
