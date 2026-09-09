import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { create, toBinary } from "@bufbuild/protobuf";
import { PresentationPreviewSceneSchema, PresentationElementSchema, PresentationSlideSchema, PresentationImageTransformSchema } from "../src/generated/office_kit/artifact/v1/office_artifact_pb.js";
import { createPpjSceneView, scenePoints, sceneDegrees, sceneOpacity, sceneFontPoints } from "../src/ppj/preview-scene-view.mjs";

const bytes = text => new TextEncoder().encode(text);
const sha = data => createHash("sha256").update(data).digest("hex");
const content = PresentationElementSchema.fields.filter(field => field.oneof?.localName === "content");
const registry = JSON.parse(readFileSync(new URL("../src/ppj/capability-registry.json", import.meta.url)));
assert.equal(registry.previewScene.adapter, "src/ppj/preview-scene-view.mjs");
assert.equal(registry.previewScene.adapterTest, "test/ppj-preview-scene-view.mjs");
assert.equal(registry.previewScene.contentSource, PresentationElementSchema.typeName + ".content");
function fixture(change = () => {}) {
  // Typed, synthetic scene evidence; not image/PPTX or visual acceptance.
  const file = bytes("candidate"), programJson = bytes('{"pages":[]}'), data = bytes("asset");
  const scene = create(PresentationPreviewSceneSchema, {
    version: 1, origin: 1, programSha256: sha(programJson), candidateSha256: sha(file),
    presentation: { slideWidthEmu: 1270000n, slideHeightEmu: 635000n, slides: [{ id: "native-slide", elements: content.map((field, i) => ({
      id: `native-${i}`, hidden: i === 0 ? false : undefined,
      content: { case: field.localName, value: {
        leftEmu: -12700n, topEmu: 25400n, widthEmu: 127000n, heightEmu: 63500n,
      } },
    })) }] },
    assets: [{ nativeId: "native-asset", contentType: "image/png", sha256: sha(data) }],
  });
  const nodes = scene.presentation.slides[0].elements;
  const body = kind => nodes.find(node => node.content.case === kind).content.value;
  body("shape").text = "same paragraph";
  body("shape").textBody = create(content.find(field => field.localName === "shape").message.fields.find(field => field.localName === "textBody").message, {
    paragraphs: [{ runs: [{ content: { case: "text", value: "A" }, fontSizePoints: 12.5 }, { content: { case: "text", value: "B" }, bold: false }] }],
  });
  body("image").assetId = "native-asset";
  body("image").opacityThousandthPercent = 0;
  body("image").transform = create(PresentationImageTransformSchema, { rotationAngle60000: -5400000, flipHorizontal: false, flipVertical: true });
  body("connector").startXEmu = 127000n;
  body("connector").startYEmu = 25400n;
  body("connector").endXEmu = -12700n;
  body("connector").endYEmu = 63500n;
  body("connector").startTargetId = "native-0";
  body("connector").startConnectionSiteIndex = 0;
  body("group").childWidthEmu = 0n;
  body("group").childHeightEmu = 12700n;
  body("group").children.push(create(PresentationElementSchema, { id: "repeat-child", content: { case: "shape", value: {
    geometry: "rect", leftEmu: 6350n, topEmu: 0n, widthEmu: 12700n, heightEmu: 12700n,
  } } }));
  body("diagram").drawing = create(content.find(field => field.localName === "group").message, {
    children: [{ id: "diagram-child", content: { case: "shape", value: { geometry: "ellipse" } } }],
  });
  body("chart").series.push(create(content.find(field => field.localName === "chart").message.fields.find(field => field.localName === "series").message, {
    values: [1, 0, 0], missingValueIndexes: [1], xValues: [8, 3, 9],
  }));
  function bindings(elements, parent) {
    elements.forEach((element, i) => {
      const scenePath = `${parent}[${i}]`, nested = parent.includes("children");
      scene.bindings.push(create(PresentationPreviewSceneSchema.fields.find(field => field.localName === "bindings").message, { pageId: "semantic-page", semanticId: nested ? "component-owner" : `semantic-${i}`, nativeId: element.id,
        programPath: nested ? "$.components[0].elements[0]" : `$.pages[0].elements[${i}]`, scenePath,
        attribution: nested ? 2 : 1, zOrder: i, componentId: nested ? "component" : "", instanceId: nested ? "instance" : "" }));
      if (element.content.case === "group") bindings(element.content.value.children, `${scenePath}.group.children`);
      if (element.content.case === "diagram") bindings(element.content.value.drawing.children, `${scenePath}.diagram.drawing.children`);
    });
  }
  bindings(nodes, "$.presentation.slides[0].elements");
  change(scene, body);
  scene.sha256 = sha(toBinary(PresentationPreviewSceneSchema, { ...scene, sha256: "" }));
  return { file, programJson, programSha256: sha(programJson), outputSha256: sha(file), sourceBound: false,
    previewScene: scene, assets: [{ id: "different-ppj-id", mimeType: "image/png", sha256: sha(data), data }] };
}

const receipt = fixture(), before = toBinary(PresentationPreviewSceneSchema, receipt.previewScene);
const view = createPpjSceneView(receipt), page = view.pages[0];
const byKind = kind => page.nodes.find(node => node.kind === kind);
assert.deepEqual(page.nodes.map(node => node.kind), content.map(field => field.localName));
assert.equal(content.length, 9, "Update the intentional coverage fixture when native content vocabulary grows.");
assert.deepEqual(view.canvas, { width: 100, height: 50 });
assert.equal(view.native, receipt.previewScene.presentation);
assert.equal(view.paintAssessment, "unassessed");
assert.deepEqual(view.diagnostics, []);
assert.equal(page.pageId, "semantic-page");
assert.equal(page.nativeId, "native-slide");
for (const node of page.nodes) {
  assert.equal(node.native, node.element.content.value);
  assert.equal(node.binding.nativeId, node.nativeId);
  assert.ok(Object.isFrozen(node));
  assert.ok(node.frame);
}
assert.equal(byKind("shape").hidden, false);
assert.equal(byKind("image").hidden, undefined);
assert.deepEqual(byKind("shape").frame, { x: -1, y: 2, width: 10, height: 5 });
assert.equal(byKind("shape").native.textBody.paragraphs[0].runs.length, 2);
assert.equal(byKind("shape").native.textBody.paragraphs[0].runs[0].fontSizePoints, 12.5);
assert.equal(byKind("shape").native.textBody.paragraphs[0].runs[1].bold, false);
assert.equal(byKind("shape").native.textBody.paragraphs[0].runs[1].content.value, "B");
assert.equal(byKind("image").transform.rotation, -90);
assert.equal(byKind("image").transform.flipH, false);
assert.equal(byKind("image").transform.flipV, true);
assert.equal(byKind("image").native.opacityThousandthPercent, 0);
assert.equal(byKind("group").childFrame.width, 0, "Zero child extent must not become the outer frame.");
assert.equal(byKind("group").children[0].frame.x, .5, "Keep child coordinate space; do not apply a second layout.");
assert.equal(byKind("group").children[0].semanticId, "component-owner");
assert.equal(byKind("group").children[0].binding.attribution, 2);
assert.equal(byKind("diagram").drawing.children[0].native.geometry, "ellipse");
assert.deepEqual(byKind("connector").endpoints, { start: { x: 10, y: 2 }, end: { x: -1, y: 5 } });
assert.equal(byKind("connector").native.startTargetId, "native-0");
assert.equal(byKind("connector").native.startConnectionSiteIndex, 0);
assert.deepEqual(byKind("chart").native.series[0].values, [1, 0, 0]);
assert.deepEqual(byKind("chart").native.series[0].missingValueIndexes, [1]);
assert.deepEqual(byKind("chart").native.series[0].xValues, [8, 3, 9]);
assert.equal(view.asset("native-asset").data, receipt.assets[0].data);
assert.equal(view.asset("different-ppj-id"), undefined);
assert.ok(Object.isFrozen(view.assets));
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, receipt.previewScene), before);

assert.equal(scenePoints(0n), 0);
assert.equal(scenePoints(127n), .01);
assert.equal(scenePoints(undefined), undefined);
assert.equal(sceneDegrees(0), 0);
assert.equal(sceneDegrees(undefined), undefined);
assert.equal(sceneOpacity(50000), .5);
assert.equal(sceneOpacity(0), 0);
assert.equal(sceneOpacity(undefined), undefined);
assert.equal(sceneFontPoints(12.5), 12.5);
for (const bad of [null, false, "12700", NaN, Infinity, .5, Number.MAX_SAFE_INTEGER + 1, 2n ** 63n])
  assert.throws(() => scenePoints(bad), { code: "preview.scene.units" });

const unknown = createPpjSceneView(fixture((scene, body) => {
  body("group").children[0].content.value.futureVisual = { brightness: 0 };
  body("chart").series[0].$unknown = [{ no: 999, wireType: 0, data: Uint8Array.of(1) }];
  body("chart").type = 999;
  scene.presentation.slides.push(create(PresentationSlideSchema));
}));
assert.equal(unknown.pages[1].pageId, undefined, "Empty native page does not invent a semantic identity.");
assert.ok(unknown.diagnostics.some(d => d.reason === "preview.scene.unmapped-page" && d.scenePath === "$.presentation.slides[1]"));
assert.ok(unknown.diagnostics.some(d => d.reason === "preview.scene.unknown-field" && d.path === "$.components[0].elements[0]" && d.scenePath.endsWith(".shape.futureVisual")));
assert.ok(unknown.diagnostics.some(d => d.reason === "preview.scene.unknown-wire-field" && d.scenePath.endsWith("series[0].$unknown[999]")));
assert.ok(unknown.diagnostics.some(d => d.reason === "preview.scene.unknown-enum"));
assert.ok(unknown.diagnostics.every(d => d.status === "partial" && d.severity === "warning"));
assert.equal(unknown.pages[0].nodes.find(n => n.kind === "group").children[0].native.futureVisual.brightness, 0);
const unknownContent = createPpjSceneView(fixture(scene => {
  const node = scene.presentation.slides[0].elements[0];
  node.content = { case: undefined };
  node.$unknown = [{ no: 99, wireType: 2, data: Uint8Array.of(0) }];
}));
assert.equal(unknownContent.pages[0].nodes[0].kind, "unknown");
assert.equal(unknownContent.pages[0].nodes[0].native, undefined);
assert.ok(unknownContent.diagnostics.some(d => d.reason === "preview.scene.unknown-content"));
const unknownUnion = createPpjSceneView(fixture((scene, body) => {
  body("shape").textBody.paragraphs[0].runs[0].highlight = { case: "futureColor", value: "scheme" };
}));
assert.ok(unknownUnion.diagnostics.some(d => d.reason === "preview.scene.unknown-oneof"));

const missing = fixture(); delete missing.previewScene;
assert.throws(() => createPpjSceneView(missing), { code: "preview.scene.missing" });
const corrupt = fixture(); corrupt.assets[0].data[0] ^= 1;
assert.throws(() => createPpjSceneView(corrupt), { code: "preview.scene.asset-mismatch" });

// A fresh process, not an already-populated module cache, establishes leaf
// loading: root must not load the adapter, and the adapter needs no render host.
for (const mode of ["root", "adapter"]) {
const lazy = spawnSync(process.execPath, ["--input-type=module", "-e", `
  import assert from "node:assert/strict";
  import { registerHooks } from "node:module";
  const root = ${mode === "root"};
  const hooks = registerHooks({ resolve(specifier, context, next) {
    assert.doesNotMatch(specifier, /^(sharp|canvas|mupdf|playwright)$|\\/presentation\\//);
    if (!root) assert.doesNotMatch(specifier, /^jszip$|office-kit-native-client|codecs\\/office-kit-runtime/);
    if (root) assert.doesNotMatch(specifier, /preview-scene|generated\\/office_kit/);
    return next(specifier, context);
  } });
  if (root) await import("./src/index.mjs");
  else {
  const { createPpjSceneView, scenePoints } = await import("./src/ppj/preview-scene-view.mjs");
  assert.equal(scenePoints(12700n), 1);
  const { fromBinary } = await import("@bufbuild/protobuf");
  const { PresentationPreviewSceneSchema } = await import("./src/generated/office_kit/artifact/v1/office_artifact_pb.js");
  const scene = fromBinary(PresentationPreviewSceneSchema, Buffer.from(${JSON.stringify(Buffer.from(before).toString("base64"))}, "base64"));
  const adapted = createPpjSceneView({ previewScene: scene, file: Buffer.from("candidate"), programJson: Buffer.from('{"pages":[]}'),
    programSha256: scene.programSha256, outputSha256: scene.candidateSha256, sourceBound: false,
    assets: [{ mimeType: "image/png", sha256: ${JSON.stringify(receipt.assets[0].sha256)}, data: Buffer.from("asset") }] });
  assert.equal(adapted.pages[0].nodes.length, ${content.length});
  }
  hooks.deregister();
`], { cwd: new URL("../", import.meta.url), encoding: "utf8" });
assert.equal(lazy.status, 0, lazy.stderr || lazy.stdout);
}
console.log("ppj native scene view ok: all content cases, exact fields, units, ownership, assets, unknown descendants and leaf imports");
