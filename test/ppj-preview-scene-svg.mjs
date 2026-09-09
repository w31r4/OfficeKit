import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { create, clone, toBinary } from "@bufbuild/protobuf";
import { PresentationPreviewSceneSchema, PresentationElementSchema, PresentationTextParagraphSchema,
  PresentationCustomGeometryPathSchema, SpreadsheetChartMarkerArtifactSchema,
  SpreadsheetChartSeriesArtifactSchema, SpreadsheetChartPointStyleArtifactSchema,
  SpreadsheetChartSurfaceFillSchema, SpreadsheetChartAxisArtifactSchema } from "../src/generated/office_kit/artifact/v1/office_artifact_pb.js";
import { nativePathData, paintPpjSceneSvg } from "../src/ppj/preview-scene-svg.mjs";

const sha = data => createHash("sha256").update(data).digest("hex");
const emu = value => BigInt(Math.round(value * 12700));
const frame = (x, y, width, height) => ({ leftEmu: emu(x), topEmu: emu(y), widthEmu: emu(width), heightEmu: emu(height) });
const child = (id, kind, value, hidden) => create(PresentationElementSchema, { id, hidden, content: { case: kind, value } });
const content = new Map(PresentationElementSchema.fields.filter(f => f.oneof?.localName === "content").map(f => [f.localName, f.message]));
const path = data => create(PresentationCustomGeometryPathSchema, data);
const command = (kind, value) => ({ command: { case: kind, value } });
const polygon = path({ width: 100n, height: 100n, stroke: false, commands: [
  command("moveTo", { x: 0n, y: 0n }), command("lineTo", { x: 100n, y: 0n }),
  command("cubicBezierTo", { control1: { x: 100n, y: 20n }, control2: { x: 20n, y: 100n }, end: { x: 0n, y: 100n } }),
  command("quadraticBezierTo", { control: { x: 50n, y: 50n }, end: { x: 0n, y: 0n } }), command("close", true),
] });
assert.equal(nativePathData(polygon, { x: 10, y: 20, width: 200, height: 100 }), "M 10 20 L 210 20 C 210 40 50 120 10 120 Q 110 70 10 20 Z");
const noViewport = path({ commands: [command("moveTo", { x: 12700n, y: 25400n }), command("lineTo", { x: 38100n, y: 0n })] });
assert.equal(nativePathData(noViewport, { x: 10, y: 20, width: 999, height: 999 }), "M 11 22 L 13 20");
const arcCommand = (widthRadius, heightRadius, startAngle, sweepAngle) => command("arcTo", {
  widthRadius: BigInt(widthRadius), heightRadius: BigInt(heightRadius),
  startAngle: startAngle * 60000, sweepAngle: sweepAngle * 60000,
});
const quarterArc = path({ width: 100n, height: 100n, commands: [
  command("moveTo", { x: 50n, y: 0n }), arcCommand(50, 50, 0, 90),
] });
assert.equal(nativePathData(quarterArc, { x: 10, y: 20, width: 200, height: 100 }), "M 110 20 A 100 50 0 0 1 10 70");
const reverseArc = path({ width: 100n, height: 100n, commands: [
  command("moveTo", { x: 50n, y: 0n }), arcCommand(50, 50, 0, -90),
] });
assert.equal(nativePathData(reverseArc, { x: 10, y: 20, width: 200, height: 100 }), "M 110 20 A 100 50 0 0 0 10 -30");
const anisotropicArc = path({ width: 100n, height: 100n, commands: [
  command("moveTo", { x: 50n, y: 0n }), arcCommand(50, 25, 45, 90),
] });
assert.equal(nativePathData(anisotropicArc, { x: 10, y: 20, width: 200, height: 100 }), "M 110 20 A 100 25 0 0 1 20.5572809000084 20");
const fullArc = path({ width: 100n, height: 100n, commands: [
  command("moveTo", { x: 50n, y: 0n }), arcCommand(50, 50, 0, 360),
] });
assert.equal(nativePathData(fullArc, { x: 10, y: 20, width: 200, height: 100 }), "M 110 20 A 100 50 0 0 1 -90 20 A 100 50 0 0 1 110 20");
const closeThenArc = path({ width: 100n, height: 100n, commands: [
  command("moveTo", { x: 0n, y: 0n }), command("lineTo", { x: 100n, y: 0n }), command("close", true),
  arcCommand(50, 50, 180, 90),
] });
assert.equal(nativePathData(closeThenArc, { x: 10, y: 20, width: 200, height: 100 }), "M 10 20 L 210 20 Z A 100 50 0 0 1 110 -30");
const arcPath = path({ width: 100n, height: 100n, stroke: false, commands: [
  command("moveTo", { x: 50n, y: 0n }), arcCommand(50, 25, 45, 90),
] });
for (const bad of [
  path({ commands: [command("lineTo", { x: 0n, y: 0n })] }),
  path({ commands: [command("moveTo", { xReference: "guide" })] }),
  path({ commands: [command("moveTo", { x: 2n ** 63n })] }),
  path({ commands: [command("moveTo", {}), command("arcTo", {})] }),
  path({ width: 100n, height: 100n, commands: [command("moveTo", { x: 50n, y: 0n }), arcCommand(0, 50, 0, 90)] }),
  path({ width: 100n, height: 100n, commands: [command("moveTo", { x: 50n, y: 0n }), arcCommand(50, 50, 0, 0)] }),
  path({ width: 100n, height: 100n, commands: [command("moveTo", { x: 50n, y: 0n }), arcCommand(50, 50, 0, 90), command("close", false)] }),
  path({ width: 100n, height: 100n, commands: [command("moveTo", { x: 50n, y: 0n }), command("arcTo", { widthRadius: 50n, heightRadius: 50n, startAngle: 0, sweepAngle: 90, widthRadiusReference: "guide" })] }),
]) assert.throws(() => nativePathData(bad, { x: 0, y: 0, width: 100, height: 100 }));

// Synthetic native receipts isolate drawing semantics; real compiler coverage
// lives in the explicitly selected NativeAOT integration test.
function fixture(edit = () => {}, program = { pages: [] }, editBindings = () => {}) {
  const file = Buffer.from("candidate"), programJson = Buffer.from(JSON.stringify(program));
  const data = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLttAAAAABJRU5ErkJggg==", "base64");
  const nodes = [
    child("box", "shape", { ...frame(10, 40, 200, 100), geometry: "textbox", fillRgb: "114477", lineRgb: "AA2200", lineWidthEmu: emu(2),
      fillOpacityThousandthPercent: 0, text: "WRONG FLATTENED TEXT", textBody: { paragraphs: [{ alignment: "center",
        defaultRunStyle: { case: "defaultRunProperties", value: { fontSizePoints: 16, bold: true, color: { case: "colorRgb", value: "123456" } } },
        runs: [{ content: { case: "text", value: "A & " }, bold: false }, { content: { case: "text", value: "B" }, fontSizePoints: 12.5 },
          { content: { case: "lineBreak", value: true } }, { content: { case: "text", value: "C" }, colorOpacityThousandthPercent: 0 }],
      }] } }),
    child("diamond", "shape", { ...frame(220, 40, 100, 100), geometry: "flowChartDecision", fillRgb: "33AA44", text: "Decision" }),
    child("vector", "shape", { ...frame(10, 150, 200, 100), geometry: "custom", fillRgb: "CC5500", customPaths: [clone(PresentationCustomGeometryPathSchema, polygon), clone(PresentationCustomGeometryPathSchema, arcPath)] }),
    child("photo", "image", { ...frame(330, 40, 50, 100), assetId: "native-asset", opacityThousandthPercent: 0, transform: { rotationAngle60000: -5400000, flipHorizontal: true, flipVertical: false } }),
    child("outer", "group", { ...frame(100, 280, 200, 100), childLeftEmu: emu(10), childTopEmu: emu(20), childWidthEmu: emu(100), childHeightEmu: emu(50),
      frameTransform: { rotationAngle60000: 5400000 }, children: [child("inner", "group", { ...frame(10, 20, 40, 20), childWidthEmu: emu(40), childHeightEmu: emu(20), children: [
        child("leaf", "shape", { ...frame(2, 3, 10, 5), geometry: "ellipse", fillRgb: "00AACC" }),
      ] })] }),
    child("hidden", "shape", { ...frame(400, 40, 50, 50), geometry: "rect", text: "HIDDEN TEXT" }, true),
    child("chart", "chart", { ...frame(400, 150, 150, 100), series: [{ values: [1, 0, 0, 3], missingValueIndexes: [1] }] }),
    child("grid", "table", { ...frame(20, 40, 300, 120), columnWidthsEmu: [emu(100), emu(200)], firstRow: true,
      mergeRanges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }], rows: [
        { heightEmu: emu(40), cells: [{ text: "Merged", fill: { kind: { case: "solidRgb", value: "AA5500" }, opacityThousandthPercent: 0 },
          textBody: { paragraphs: [{ runs: [{ content: { case: "text", value: "Mer" }, fontSizePoints: 12 }, { content: { case: "text", value: "ged" }, bold: true, fontSizePoints: 12 }] }] } }, {}] },
        { heightEmu: emu(80), cells: [{ text: "0", fill: { kind: { case: "solidRgb", value: "00AA44" } },
          textStyle: { fontSizePoints: 10, bold: false, color: { case: "colorRgb", value: "556677" }, colorOpacityThousandthPercent: 0 },
          borders: { right: { color: { source: { case: "rgb", value: "0000FF" } }, widthPoints: 2, opacityThousandthPercent: 50000 } } }, { text: "Wider", fill: { kind: { case: "noFill", value: true } } }] },
      ] }),
    child("directed", "connector", { connectorType: "straight", startXEmu: emu(500), startYEmu: emu(350), endXEmu: emu(350), endYEmu: emu(280),
      startTargetId: "diamond", endTargetId: "grid", startConnectionSiteIndex: 0, endConnectionSiteIndex: 3,
      lineRgb: "336699", lineWidthEmu: emu(2), lineOpacityThousandthPercent: 50000, startArrow: "diamond", endArrow: "triangle", endArrowWidth: "lg", endArrowLength: "sm" }),
  ];
  const scene = create(PresentationPreviewSceneSchema, { version: 1, origin: 1, programSha256: sha(programJson), candidateSha256: sha(file),
    presentation: { slideWidthEmu: emu(600), slideHeightEmu: emu(400), slides: [{ id: "native-page", elements: nodes }] },
    assets: [{ nativeId: "native-asset", contentType: "image/png", sha256: sha(data) }],
  });
  edit(scene);
  function bind(elements, prefix) {
    elements.forEach((element, i) => {
      const scenePath = `${prefix}[${i}]`;
      scene.bindings.push(create(PresentationPreviewSceneSchema.fields.find(f => f.localName === "bindings").message, {
        pageId: "page", semanticId: element.id === "leaf" ? "component-owner" : element.id, nativeId: element.id,
        programPath: "$.pages[0]", scenePath, attribution: 1, zOrder: i,
      }));
      if (element.content.case === "group") bind(element.content.value.children, `${scenePath}.group.children`);
      if (element.content.case === "diagram" && element.content.value.drawing) bind(element.content.value.drawing.children, `${scenePath}.diagram.drawing.children`);
    });
  }
  bind(scene.presentation.slides[0].elements, "$.presentation.slides[0].elements");
  editBindings(scene.bindings);
  scene.sha256 = sha(toBinary(PresentationPreviewSceneSchema, { ...scene, sha256: "" }));
  return { file, programJson, programSha256: sha(programJson), outputSha256: sha(file), sourceBound: scene.origin === 2, previewScene: scene,
    assets: [{ id: "public-asset", mimeType: "image/png", sha256: sha(data), data }] };
}
const receipt = fixture(), before = toBinary(PresentationPreviewSceneSchema, receipt.previewScene);
const combinedInput = { pages: [{ id: "page", elements: [{ id: "owner", type: "shape",
  frame: { x: 10, y: 40, width: 100, height: 100, rotation: 30 },
  geometry: { kind: "preset", preset: "rect" }, futurePaint: "unmapped input" }] }],
  futureGlobal: "unresolved global input" };
const combinedReceipt = fixture(scene => {
  scene.presentation.slides[0].elements = scene.presentation.slides[0].elements.slice(0, 2);
}, combinedInput, bindings => {
  for (const binding of bindings) {
    binding.semanticId = "owner";
    binding.programPath = "$.pages[0].elements[0]";
    binding.attribution = 2;
  }
});
const combinedBefore = toBinary(PresentationPreviewSceneSchema, combinedReceipt.previewScene);
const paintOnly = paintPpjSceneSvg(combinedReceipt);
const combinedAssessment = paintPpjSceneSvg(combinedReceipt, { assessInput: true });
assert.ok(!paintOnly.inputAssessment);
assert.ok(combinedAssessment.inputAssessment);
const inputTransform = combinedAssessment.inputAssessment.diagnostics.find(d => d.path.endsWith(".frame.rotation") && d.severity === "error");
assert.ok(inputTransform, "unpainted requested rotation remains factual failure");
for (const node of combinedAssessment.pages[0].assessment.children) {
  assert.ok(node.diagnostics.some(d => d.path === inputTransform.path && d.scenePath === node.scenePath && d.severity === "error"));
  assert.ok(node.diagnostics.some(d => d.path.endsWith(".futurePaint") && d.scenePath === node.scenePath));
}
assert.equal(combinedAssessment.reliability.status, "failed");
assert.match(combinedAssessment.pages[0].svg, /data-officekit-review="failed"/);
assert.match(combinedAssessment.pages[0].svg, /fill="#991B1B"/);
assert.ok(combinedAssessment.diagnostics.some(d => d.path === "$.futureGlobal" && d.scenePath === "$.presentation"));
for (const d of paintOnly.diagnostics) assert.ok(combinedAssessment.diagnostics.some(a => a.reason === d.reason && a.scenePath === d.scenePath));
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, combinedReceipt.previewScene), combinedBefore);
const inputAsset = receipt.assets[0];
const assetProgram = { assets: [{ id: "different-public-id", mimeType: inputAsset.mimeType, sha256: inputAsset.sha256.toUpperCase() }],
  pages: [{ id: "page", elements: [{ id: "photo", type: "image", asset: "different-public-id", fit: "contain",
    frame: { x: 330, y: 40, width: 50, height: 100 } }] }] };
const assetReceipt = fixture(scene => { scene.presentation.slides[0].elements = [scene.presentation.slides[0].elements[3]]; }, assetProgram,
  bindings => { bindings[0].programPath = "$.pages[0].elements[0]"; });
assert.ok(!paintPpjSceneSvg(assetReceipt, { assessInput: true }).diagnostics.some(d => d.reason === "asset-missing"),
  "different public/native IDs resolve via verified MIME/hash, including uppercase declaration digest");
function groupProfileFixture({ requestedWidth = 100, nativeWidth = 100, hidden = false } = {}) {
  return fixture(scene => {
    scene.presentation.slides[0].elements = [child("outer", "group", {
      ...frame(100, 100, 200, 100), childLeftEmu: emu(10), childTopEmu: emu(20),
      childWidthEmu: emu(nativeWidth), childHeightEmu: emu(50), children: [],
    }, hidden)];
  }, { pages: [{ id: "page", elements: [{ id: "outer", type: "group", hidden,
    frame: { x: 100, y: 100, width: 200, height: 100 },
    childFrame: { x: 10, y: 20, width: requestedWidth, height: 50 }, elements: [] }] }] },
  bindings => { bindings[0].programPath = "$.pages[0].elements[0]"; });
}
const groupReason = "preview.fact.group-coordinates-ignored";
assert.ok(!paintPpjSceneSvg(groupProfileFixture(), { assessInput: true }).diagnostics.some(d => d.reason === groupReason));
for (const options of [{ requestedWidth: 80 }, { hidden: true }, { nativeWidth: 0 }]) {
  const result = paintPpjSceneSvg(groupProfileFixture(options), { assessInput: true });
  assert.ok(result.diagnostics.some(d => d.reason === groupReason), "mismatch, hidden and failed group retain old factual rule");
  assert.equal(result.reliability.status, "failed");
}
const attributed = fixture(scene => {
  scene.presentation.futureReviewField = "unmodeled global state";
  scene.presentation.slides[0].elements = ["generated-a", "generated-b"].map(id => child(id, "shape", {
    ...frame(10, 10, 40, 40), geometry: "rect", fillRgb: "CC5500", shapeDepthEmu: 12700n, shapeExtrusionHeightEmu: 25400n,
  }));
});
for (const binding of attributed.previewScene.bindings) {
  binding.semanticId = "component-owner"; binding.programPath = "$.pages[0].elements[0]"; binding.attribution = 2;
}
attributed.previewScene.sha256 = sha(toBinary(PresentationPreviewSceneSchema, { ...attributed.previewScene, sha256: "" }));
const attributedPaint = paintPpjSceneSvg(attributed);
const remainingGeometry = attributedPaint.diagnostics.filter(d => /shape(?:DepthEmu|ExtrusionHeightEmu)$/.test(d.scenePath));
assert.equal(remainingGeometry.length, 4, "Two unpainted fields on two generated children retain four distinct addresses");
assert.equal(new Set(remainingGeometry.map(d => d.path)).size, 1);
assert.equal(new Set(remainingGeometry.map(d => d.scenePath)).size, 4);
assert.equal(attributedPaint.pages[0].id, "page");
assert.equal(attributedPaint.pages[0].assessment.scenePath, "$.presentation.slides[0]");
assert.ok(attributedPaint.pages[0].diagnostics.some(d => d.scenePath === "$.presentation.futureReviewField"));
assert.deepEqual(attributedPaint.assessment.diagnostics, attributedPaint.diagnostics);
assert.deepEqual(attributedPaint.pages[0].assessment.diagnostics, attributedPaint.pages[0].diagnostics);
assert.equal(JSON.parse(JSON.stringify(attributedPaint.assessment)).diagnostics.length, attributedPaint.diagnostics.length);
const generatedAssessments = attributedPaint.pages[0].assessment.children;
assert.equal(generatedAssessments.length, 2);
for (const [index, assessment] of generatedAssessments.entries()) {
  assert.equal(assessment.id, "component-owner");
  assert.equal(assessment.path, "$.pages[0].elements[0]");
  assert.equal(assessment.scenePath, `$.presentation.slides[0].elements[${index}]`);
  assert.equal(assessment.diagnostics.filter(d => /shape(?:DepthEmu|ExtrusionHeightEmu)$/.test(d.scenePath)).length, 2);
  assert.ok(assessment.diagnostics.every(d => d.scenePath.startsWith(`${assessment.scenePath}.`)));
}
for (const hidden of [true, false]) {
  const skipped = fixture(scene => {
    scene.presentation.slides[0].elements = [child("parent", "group", {
      ...frame(0, 0, 100, 100), childWidthEmu: 0n, childHeightEmu: emu(100),
      children: [child("unvisited", "shape", { ...frame(0, 0, 10, 10), geometry: "rect", fillRgb: "CC5500" })],
    }, hidden)];
  });
  const painted = paintPpjSceneSvg(skipped);
  const parent = painted.pages[0].assessment.children[0];
  assert.equal(parent.children.length, 1);
  const leaf = parent.children[0];
  assert.ok(leaf.diagnostics.some(d => d.reason === "preview.state.unassessed" && d.scenePath === leaf.scenePath));
  assert.notEqual(leaf.reliability.status, "passed");
  if (!hidden) assert.equal(parent.reliability.status, "failed");
}
for (const verified of [true, false]) {
  const input = fixture(scene => scene.presentation.slides[0].elements.push(child("diagram", "diagram", {
    ...frame(100, 100, 200, 100), drawingCacheVerified: verified, layout: "process",
    drawing: { ...frame(100, 100, 200, 100), childLeftEmu: emu(10), childTopEmu: emu(20), childWidthEmu: emu(100), childHeightEmu: emu(50),
      children: [child("cached-node", "shape", { ...frame(10, 20, 40, 20), geometry: "rect", fillRgb: "AA5500", text: "Cached label" })] },
  })));
  const original = toBinary(PresentationPreviewSceneSchema, input.previewScene);
  const result = paintPpjSceneSvg(input);
  if (verified) {
    assert.match(result.pages[0].svg, /data-officekit-diagram="verified-cache"/);
    assert.match(result.pages[0].svg, /translate\(100 100\) scale\(2 2\) translate\(-10 -20\)/);
    assert.match(result.pages[0].svg, /data-officekit-native-id="cached-node"/);
    assert.match(result.pages[0].svg, /Cached label/);
    assert.ok(result.diagnostics.some(d => d.scenePath.includes("diagram.drawing.children[0].shape.textBody")));
  } else {
    assert.match(result.pages[0].svg, /diagram: verified drawing unavailable/);
    assert.ok(!result.pages[0].svg.includes("Cached label"));
  }
  assert.deepEqual(toBinary(PresentationPreviewSceneSchema, input.previewScene), original);
}
for (const [dash, pattern] of [["dashed", "8 6"], ["dotted", "2 6"], ["dash-dot", "8 6 2 6"], ["dash-dot-dot", "16 6 2 6 2 6"]]) {
  const input = fixture(scene => {
    const s = scene.presentation.slides[0].elements[0].content.value;
    Object.assign(s, { lineStyle: dash, lineCap: "round", lineJoin: "bevel", lineOpacityThousandthPercent: 50000 });
  });
  const original = toBinary(PresentationPreviewSceneSchema, input.previewScene);
  const result = paintPpjSceneSvg(input);
  assert.ok(result.pages[0].svg.includes(`stroke="#AA2200" stroke-width="2" stroke-opacity="0.5" stroke-linecap="round" stroke-linejoin="bevel" stroke-dasharray="${pattern}"`));
  assert.ok(result.diagnostics.some(d => d.reason === "preview.scene.paint.dash-metrics" && d.scenePath.endsWith("shape.lineStyle")));
  assert.ok(!result.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && /\.(lineCap|lineJoin)$/.test(d.scenePath)));
  assert.deepEqual(toBinary(PresentationPreviewSceneSchema, input.previewScene), original);
}
for (const [lineWidthEmu, lineOpacityThousandthPercent] of [[0n, 100000], [emu(2), 0]]) {
  const painted = paintPpjSceneSvg(fixture(scene => Object.assign(scene.presentation.slides[0].elements[0].content.value,
    { lineWidthEmu, lineOpacityThousandthPercent })));
  assert.ok(painted.pages[0].svg.includes(`stroke="#AA2200" stroke-width="${Number(lineWidthEmu) / 12700}" stroke-opacity="${lineOpacityThousandthPercent / 100000}"`));
}
const strokelessPath = paintPpjSceneSvg(fixture(scene => {
  Object.assign(scene.presentation.slides[0].elements[2].content.value,
    { lineStyle: "dashed", lineCap: "round", lineJoin: "bevel", lineRgb: "FF0000", lineWidthEmu: emu(2) });
}));
assert.match(strokelessPath.pages[0].svg, /stroke-dasharray="8 6"><path data-officekit-path="0"[^>]*stroke="none"/);
const invalidShapeStroke = paintPpjSceneSvg(fixture(scene => {
  scene.presentation.slides[0].elements[0].content.value.lineCap = "invented";
}));
assert.equal(invalidShapeStroke.reliability.status, "failed");
for (const [margin, indent] of [[30, -10], [30, 10], [0, 0], [-5, 0]]) {
  const input = fixture(scene => {
    const p = scene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0];
    p.alignment = "left";
    p.leftMargin = { case: "marginLeftEmu", value: emu(margin) };
    p.indentation = { case: "indentEmu", value: emu(indent) };
  });
  const original = toBinary(PresentationPreviewSceneSchema, input.previewScene);
  const painted = paintPpjSceneSvg(input);
  const xs = [...painted.pages[0].svg.matchAll(/<text x="([^"]+)" y="(?:59.6|78.8)"/g)].map(m => Number(m[1]));
  assert.equal(xs.length, 2);
  assert.ok(Math.abs(xs[0] - (17.2 + margin + indent)) < 1e-9);
  assert.ok(Math.abs(xs[1] - (17.2 + margin)) < 1e-9);
  assert.deepEqual(toBinary(PresentationPreviewSceneSchema, input.previewScene), original);
  assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && /\.(marginLeftEmu|indentEmu)$/.test(d.scenePath)));
}
const centeredIndent = paintPpjSceneSvg(fixture(scene => {
  scene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0].indentation = { case: "indentEmu", value: emu(10) };
}));
assert.ok(centeredIndent.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith("indentEmu")));
const paragraphSpacing = fixture(scene => {
  const body = scene.presentation.slides[0].elements[0].content.value.textBody;
  const p = body.paragraphs[0];
  p.lineSpacing = { case: "lineSpacingPoints", value: 30 };
  p.spaceBefore = { case: "spaceBeforePoints", value: 5 };
  p.spaceAfter = { case: "spaceAfterPoints", value: 7 };
  body.paragraphs.push(clone(PresentationTextParagraphSchema, p));
});
const paragraphOriginal = toBinary(PresentationPreviewSceneSchema, paragraphSpacing.previewScene);
const paragraphPaint = paintPpjSceneSvg(paragraphSpacing);
const textYs = [...paragraphPaint.pages[0].svg.matchAll(/<text x="110" y="([^"]+)"/g)].map(m => Number(m[1]));
assert.equal(textYs.length, 4);
assert.ok(Math.abs(textYs[0] - 64.6) < 1e-9);
assert.ok(Math.abs(textYs[1] - textYs[0] - 30) < 1e-9);
assert.ok(Math.abs(textYs[2] - textYs[1] - 31.2) < 1e-9);
assert.ok(Math.abs(textYs[3] - textYs[2] - 30) < 1e-9);
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, paragraphSpacing.previewScene), paragraphOriginal);
assert.ok(!paragraphPaint.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && /\.(lineSpacingPoints|spaceBeforePoints|spaceAfterPoints)$/.test(d.scenePath)));
const zeroParagraph = paintPpjSceneSvg(fixture(scene => {
  const p = scene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0];
  p.spaceBefore = { case: "spaceBeforePoints", value: 0 };
  p.spaceAfter = { case: "spaceAfterPoints", value: 0 };
  p.lineSpacing = { case: "lineSpacingMultiplier", value: 2 };
}));
assert.match(zeroParagraph.pages[0].svg, /<text x="110" y="59.6"/);
assert.ok(zeroParagraph.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith("lineSpacingMultiplier")));
for (const value of [-1, 0, Infinity]) {
  const result = paintPpjSceneSvg(fixture(scene => {
    scene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0].lineSpacing = { case: "lineSpacingPoints", value };
  }));
  assert.equal(result.reliability.status, "failed");
  assert.ok(result.diagnostics.some(d => d.reason === "preview.scene.paint.paragraph-spacing"));
}
for (const spacing of [-768, -2.5, 0, 3.25, 768]) {
  const input = fixture(scene => {
    const p = scene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0];
    p.defaultRunStyle.value.fontSpacingPoints = spacing;
    p.runs[1].fontSpacingPoints = 0;
  });
  const original = toBinary(PresentationPreviewSceneSchema, input.previewScene);
  const result = paintPpjSceneSvg(input);
  assert.ok(result.pages[0].svg.includes(`letter-spacing="${spacing}"`));
  assert.match(result.pages[0].svg, /letter-spacing="0"[^>]*>B<\/tspan>/);
  assert.deepEqual(toBinary(PresentationPreviewSceneSchema, input.previewScene), original);
  assert.ok(!result.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith("fontSpacingPoints")));
}
for (const value of [-769, 769, Infinity]) {
  const result = paintPpjSceneSvg(fixture(scene => {
    scene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0].runs[0].fontSpacingPoints = value;
  }));
  assert.equal(result.reliability.status, "failed");
  assert.ok(result.diagnostics.some(d => d.reason === "preview.scene.paint.text-spacing"));
}
const textFormats = fixture(scene => {
  const paragraph = scene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0];
  Object.assign(paragraph.defaultRunStyle.value, { underline: "sng", strike: "sngStrike", fontBaselinePercent: 25 });
  Object.assign(paragraph.runs[1], { underline: "none", strike: "noStrike", fontBaselinePercent: 0 });
  Object.assign(paragraph.runs[3], { fontBaselinePercent: -50 });
});
const textFormatsBefore = toBinary(PresentationPreviewSceneSchema, textFormats.previewScene);
const textFormatsPaint = paintPpjSceneSvg(textFormats);
assert.match(textFormatsPaint.pages[0].svg, /<tspan text-decoration="underline line-through" dy="-4"[^>]*>A &amp; <\/tspan><tspan text-decoration="none" dy="4"[^>]*>B<\/tspan>/);
assert.match(textFormatsPaint.pages[0].svg, /<tspan text-decoration="underline line-through" dy="8"[^>]*>C<\/tspan>/);
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, textFormats.previewScene), textFormatsBefore);
assert.ok(!textFormatsPaint.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && /\.(underline|strike|fontBaselinePercent)$/.test(d.scenePath)));
for (const [field, value, reason] of [["underline", "wavyDbl", "text-decoration"], ["strike", "dblStrike", "text-decoration"], ["fontBaselinePercent", 401, "text-baseline"]]) {
  const result = paintPpjSceneSvg(fixture(scene => {
    scene.presentation.slides[0].elements[0].content.value.textBody.paragraphs[0].runs[0][field] = value;
  }));
  assert.equal(result.reliability.status, "failed");
  assert.ok(result.diagnostics.some(d => d.reason === `preview.scene.paint.${reason}` && d.scenePath.endsWith(field)));
}
for (const [width, alpha] of [[2, 50000], [0, 0]]) {
  const result = paintPpjSceneSvg(fixture(scene => {
    const photo = scene.presentation.slides[0].elements.find(e => e.content.case === "image").content.value;
    photo.maskPreset = "ellipse";
    photo.border = create(content.get("image").fields.find(f => f.localName === "border").message,
      { colorRgb: "FF00FF", widthEmu: emu(width), opacityThousandthPercent: alpha, style: "dashed", cap: "round", join: "bevel" });
  }));
  assert.ok(result.pages[0].svg.includes(`data-officekit-image-border="true" fill="none" stroke="#FF00FF" stroke-width="${width}" stroke-opacity="${alpha / 100000}"`));
  assert.match(result.pages[0].svg, /stroke-linecap="round" stroke-linejoin="bevel"[^>]*><ellipse/);
}
const schemeBorder = paintPpjSceneSvg(fixture(scene => {
  scene.presentation.slides[0].elements.find(e => e.content.case === "image").content.value.border =
    create(content.get("image").fields.find(f => f.localName === "border").message, { colorScheme: "accent1", widthEmu: emu(2) });
}));
assert.equal(schemeBorder.reliability.status, "failed");
assert.ok(schemeBorder.diagnostics.some(d => d.reason === "preview.scene.paint.image-border"));
for (const preset of ["ellipse", "diamond", "rect", "triangle"]) {
  const input = fixture(scene => {
    scene.presentation.slides[0].elements.find(e => e.content.case === "image").content.value.maskPreset = preset;
  });
  const result = paintPpjSceneSvg(input);
  if (preset === "triangle") {
    assert.equal(result.reliability.status, "failed");
    assert.match(result.pages[0].svg, /Image mask unavailable/);
  } else if (preset !== "rect") {
    assert.match(result.pages[0].svg, /clipPathUnits="userSpaceOnUse"/);
    assert.match(result.pages[0].svg, /clip-path="url\(#officekit-image-mask-0\)"/);
    assert.equal(result.pages[0].svg, paintPpjSceneSvg(input).pages[0].svg);
  }
}
const customMask = paintPpjSceneSvg(fixture(scene => {
  scene.presentation.slides[0].elements.find(e => e.content.case === "image").content.value.customMaskPaths = [polygon];
}));
assert.match(customMask.pages[0].svg, /<clipPath[^>]*><path d="M/);
const badMask = paintPpjSceneSvg(fixture(scene => {
  const photo = scene.presentation.slides[0].elements.find(e => e.content.case === "image").content.value;
  photo.maskPreset = "ellipse"; photo.maskPresetAdjustments = [10000];
}));
assert.equal(badMask.reliability.status, "failed");
assert.ok(badMask.diagnostics.some(d => d.reason === "preview.scene.paint.image-mask"));
for (const [adjustments, radius] of [[[], 8.3335], [[0], 0], [[25000], 12.5], [[50000], 25], [[75000], 25], [[-100], 0]]) {
  const input = fixture(scene => {
    const photo = scene.presentation.slides[0].elements.find(e => e.content.case === "image").content.value;
    photo.maskPreset = "roundRect"; photo.maskPresetAdjustments = adjustments;
    const shape = scene.presentation.slides[0].elements[0].content.value;
    shape.geometry = "roundRect"; shape.presetAdjustments = adjustments;
  });
  const result = paintPpjSceneSvg(input);
  assert.ok(result.pages[0].svg.includes(`width="50" height="100" rx="${radius}" ry="${radius}"`));
  assert.ok(result.pages[0].svg.includes(`width="200" height="100" rx="${radius * 2}" ry="${radius * 2}"`));
  assert.ok(!result.diagnostics.some(d => d.scenePath?.endsWith("presetAdjustments") || d.reason === "preview.scene.paint.image-mask"));
}
const painted = paintPpjSceneSvg(receipt), svg = painted.pages[0].svg;
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, receipt.previewScene), before);
assert.equal(painted.scene, receipt.previewScene);
assert.match(svg, /<rect x="10" y="40" width="200" height="100" fill="#114477" fill-opacity="0"/);
assert.match(svg, /stroke-width="2"/);
assert.match(svg, /<text x="110"[^>]*text-anchor="middle"[^>]*><tspan[^>]*font-weight="normal"[^>]*>A &amp; <\/tspan><tspan[^>]*font-size="12.5"[^>]*font-weight="bold"[^>]*>B<\/tspan><\/text>/);
assert.doesNotMatch(svg, /WRONG FLATTENED|HIDDEN TEXT/);
assert.match(svg, /d="M 270 40 L 320 90 L 270 140 L 220 90 Z"/);
assert.match(svg, /data-officekit-path="0" d="M 10 150 L 210 150 C 210 170 50 250 10 250 Q 110 200 10 150 Z" stroke="none"/);
assert.match(svg, /data-officekit-path="1" d="M 110 150 A 100 25 0 0 1 20.5572809000084 150" stroke="none"/);
assert.match(svg, /translate\(355 90\) rotate\(-90\) scale\(-1 1\) translate\(-355 -90\)/);
assert.match(svg, /<image x="330" y="40" width="50" height="100"[^>]*preserveAspectRatio="none" opacity="0"/);
assert.equal(Buffer.from(svg.match(/href="data:image\/png;base64,([^"]+)"/)[1], "base64").compare(receipt.assets[0].data), 0);
for (const [crop, position] of [
  [{ leftThousandthPercent: 50000 }, 'x="-50" y="0" width="100" height="100"'],
  [{ topThousandthPercent: 50000 }, 'x="0" y="-100" width="50" height="200"'],
  [{ leftThousandthPercent: -50000, rightThousandthPercent: -50000 }, 'x="12.5" y="0" width="25" height="100"'],
  [{}, 'x="0" y="0" width="50" height="100"'],
]) {
  const input = fixture(scene => {
    const photo = scene.presentation.slides[0].elements.find(e => e.content.case === "image").content.value;
    photo.crop = create(content.get("image").fields.find(f => f.localName === "crop").message, crop);
  });
  const snapshot = toBinary(PresentationPreviewSceneSchema, input.previewScene);
  const result = paintPpjSceneSvg(input);
  assert.ok(result.pages[0].svg.includes(`<svg x="330" y="40" width="50" height="100" viewBox="0 0 50 100" overflow="hidden"><image ${position}`));
  assert.ok(!result.diagnostics.some(d => d.scenePath?.endsWith("image.crop")));
  assert.deepEqual(toBinary(PresentationPreviewSceneSchema, input.previewScene), snapshot);
}
for (const crop of [{ leftThousandthPercent: 100000 }, { topThousandthPercent: 60000, bottomThousandthPercent: 40000 }, { rightThousandthPercent: -100001 }]) {
  const result = paintPpjSceneSvg(fixture(scene => {
    const photo = scene.presentation.slides[0].elements.find(e => e.content.case === "image").content.value;
    photo.crop = create(content.get("image").fields.find(f => f.localName === "crop").message, crop);
  }));
  assert.equal(result.reliability.status, "failed");
  assert.ok(result.diagnostics.some(d => d.reason === "preview.scene.paint.image-crop"));
}
assert.match(svg, /translate\(100 280\) scale\(2 2\) translate\(-10 -20\)/);
assert.match(svg, /translate\(10 20\) scale\(1 1\) translate\(0 0\)/);
assert.match(svg, /data-officekit-native-id="leaf"[^>]*data-officekit-id="component-owner"/);
assert.match(svg, /<ellipse cx="7" cy="5.5" rx="5" ry="2.5"/);
assert.match(svg, /data-officekit-native-id="hidden"[^>]*display="none"/);
assert.match(svg, /data-officekit-review="requires-review"/);
assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-layout"));
assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.chart-type"));
assert.deepEqual(painted.scene.presentation.slides[0].elements[6].content.value.series[0].missingValueIndexes, [1]);
assert.deepEqual(painted.scene.presentation.slides[0].elements[6].content.value.series[0].values, [1, 0, 0, 3]);
assert.match(svg, /data-officekit-table-cell="0:0" data-officekit-row-span="1" data-officekit-column-span="2"><rect x="20" y="40" width="300" height="40" fill="#AA5500" fill-opacity="0"/);
assert.doesNotMatch(svg, /data-officekit-table-cell="0:1"/);
assert.match(svg, /data-officekit-table-cell="1:0"[^>]*><rect x="20" y="80" width="100" height="80" fill="#00AA44"/);
assert.match(svg, /data-officekit-table-cell="1:1"[^>]*><rect x="120" y="80" width="200" height="80" fill="none"/);
assert.match(svg, /data-officekit-cell-border="right" x1="120" y1="80" x2="120" y2="160" stroke="#0000FF" stroke-width="2" stroke-opacity="0.5"/);
assert.match(svg, /font-size="10" font-weight="normal"[^>]*fill="#556677" fill-opacity="0">0<\/tspan>/);
assert.match(svg, />Mer<\/tspan><tspan[^>]*font-weight="bold"[^>]*>ged<\/tspan>/);
assert.match(svg, /data-officekit-connector="straight" d="M 500 350 L 350 280"/);
assert.match(svg, /data-officekit-start-target="diamond" data-officekit-start-site="0" data-officekit-end-target="grid" data-officekit-end-site="3"/);
assert.match(svg, /data-officekit-arrow="end" data-officekit-arrow-kind="triangle" transform="translate\(350 280\) rotate\(-/);
assert.match(svg, /data-officekit-arrow="start" data-officekit-arrow-kind="diamond" transform="translate\(500 350\) rotate\([0-9]/);
assert.match(svg, /d="M 0 0 L -4 -5 L -4 5 Z"/);
assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.content" && ["table", "connector"].includes(d.valueSummary)));

const changed = paintPpjSceneSvg(fixture(scene => {
  const table = scene.presentation.slides[0].elements[7].content.value;
  table.widthEmu = emu(600); table.heightEmu = emu(240); // Native frame scaling retains the unequal grid.
  const edge = scene.presentation.slides[0].elements[8].content.value;
  edge.connectorType = "elbow"; edge.endXEmu = emu(200); edge.endYEmu = emu(100); edge.lineOpacityThousandthPercent = 0;
}));
assert.match(changed.pages[0].svg, /data-officekit-table-cell="1:0"[^>]*><rect x="20" y="120" width="200" height="160"/);
assert.match(changed.pages[0].svg, /data-officekit-connector="elbow" d="M 500 350 L 350 350 L 350 100 L 200 100"/);
assert.match(changed.pages[0].svg, /data-officekit-arrow="end"[^>]*translate\(200 100\) rotate\(180\)[^>]*opacity="0"/);
for (const bad of [
  table => { table.columnWidthsEmu[0] = 0n; },
  table => { table.mergeRanges[0].endRow = 20; },
  table => { table.mergeRanges.push(table.mergeRanges[0]); },
  table => { table.rows[0].cells[1].text = "must not silently disappear"; },
]) {
  const failure = paintPpjSceneSvg(fixture(scene => bad(scene.presentation.slides[0].elements[7].content.value)));
  assert.equal(failure.reliability.status, "failed");
  assert.ok(failure.diagnostics.some(d => d.id === "grid" && d.reason === "preview.scene.paint.failed"));
  assert.doesNotMatch(failure.pages[0].svg, /data-officekit-table-cell=/);
}
const adjusted = paintPpjSceneSvg(fixture(scene => {
  const edge = scene.presentation.slides[0].elements[8].content.value;
  edge.connectorType = "elbow"; edge.bendAdjustment = 0;
}));
assert.match(adjusted.pages[0].svg, /data-officekit-connector="elbow" d="M 500 350 L 500 350 L 500 280 L 350 280"/);
assert.ok(!adjusted.diagnostics.some(d => d.scenePath?.endsWith("connector.bendAdjustment")));
for (const [bend, mid] of [[25000, 462.5], [50000, 425], [100000, 350], [-50000, 575], [150000, 275]]) {
  const result = paintPpjSceneSvg(fixture(scene => {
    const edge = scene.presentation.slides[0].elements[8].content.value;
    edge.connectorType = "elbow"; edge.bendAdjustment = bend;
  }));
  assert.ok(result.pages[0].svg.includes(`d="M 500 350 L ${mid} 350 L ${mid} 280 L 350 280"`));
  assert.ok(!result.diagnostics.some(d => d.scenePath?.endsWith("connector.bendAdjustment")));
}
const straightBend = paintPpjSceneSvg(fixture(scene => {
  scene.presentation.slides[0].elements[8].content.value.bendAdjustment = 50000;
}));
assert.equal(straightBend.reliability.status, "failed");
assert.ok(straightBend.diagnostics.some(d => d.reason === "preview.scene.paint.connector-bend" && d.status === "unavailable"));

const curved = paintPpjSceneSvg(fixture(scene => { scene.presentation.slides[0].elements[8].content.value.connectorType = "curved"; }));
assert.equal(curved.reliability.status, "failed");
assert.match(curved.pages[0].svg, /curved: route unavailable/);
assert.doesNotMatch(curved.pages[0].svg, /data-officekit-connector="straight"/);
const imported = paintPpjSceneSvg(fixture(scene => {
  scene.origin = 2;
  scene.presentation.slides[0].elements[7].content.value.rows[1].cells[0].fill = undefined;
}));
assert.ok(imported.diagnostics.some(d => d.reason === "preview.scene.paint.table-inherited-fill" && d.scenePath.endsWith("table.rows[1].cells[0].fill")));
assert.match(imported.pages[0].svg, /data-officekit-table-cell="1:0"[^>]*><rect[^>]*fill="none"/);
const vertical = paintPpjSceneSvg(fixture(scene => {
  const t = scene.presentation.slides[0].elements[7].content.value;
  t.mergeRanges[0].endRow = 1; t.mergeRanges[0].endColumn = 0;
  t.rows[1].cells[0].text = ""; t.rows[1].cells[0].textStyle = undefined;
}));
assert.match(vertical.pages[0].svg, /data-officekit-table-cell="0:0" data-officekit-row-span="2" data-officekit-column-span="1"><rect x="20" y="40" width="100" height="120"/);
assert.doesNotMatch(vertical.pages[0].svg, /data-officekit-table-cell="1:0"/);
assert.ok(vertical.diagnostics.some(d => d.reason === "preview.scene.paint.merged-cell-style"));

function scatterFixture(edit = () => {}) {
  return fixture(scene => {
    const c = child("scatter", "chart", { ...frame(100, 100, 500, 300), type: 6,
      scatterStyle: "lineWithMarkers", xAxis: { minimum: 0, maximum: 100 }, yAxis: { minimum: 0, maximum: 10 },
      series: [{ name: "XY", xValues: [0, 10, 50, 100], values: [0, 2, 0, 10], missingValueIndexes: [2],
        line: { color: { source: { case: "rgb", value: "114477" } }, widthPoints: 2 },
        marker: { symbol: 3, size: 8, fill: { source: { case: "rgb", value: "FF0000" } } } }],
    });
    edit(c.content.value); scene.presentation.slides[0].elements = [c];
  });
}
const scatterInput = scatterFixture(), scatterBytes = toBinary(PresentationPreviewSceneSchema, scatterInput.previewScene);
const scatter = paintPpjSceneSvg(scatterInput), scatterSvg = scatter.pages[0].svg;
assert.equal(scatter.reliability.status, "requires-review");
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, scatterInput.previewScene), scatterBytes);
assert.match(scatterSvg, /data-officekit-chart="scatter"/);
assert.match(scatterSvg, /data-officekit-line-segment="0:1" d="M 150 355 L 190 313"/);
assert.match(scatterSvg, /data-officekit-point="0" data-officekit-x-value="0" data-officekit-value="0"/);
assert.match(scatterSvg, /data-officekit-missing-point="2" data-officekit-x-value="50"/);
assert.doesNotMatch(scatterSvg, /data-officekit-point="2"|data-officekit-line-segment="0:3"/);
assert.match(scatterSvg, /circle cx="550" cy="145" r="4"/);
const scatterMarker = paintPpjSceneSvg(scatterFixture(c => { c.scatterStyle = "marker"; })).pages[0].svg;
assert.doesNotMatch(scatterMarker, /data-officekit-line-segment=/);
assert.equal((scatterMarker.match(/<circle /g) || []).length, 3);
const scatterReverse = paintPpjSceneSvg(scatterFixture(c => { c.xAxis.reverse = c.yAxis.reverse = true; })).pages[0].svg;
assert.match(scatterReverse, /d="M 550 145 L 510 187"/);
const scatterLog = paintPpjSceneSvg(scatterFixture(c => {
  c.xAxis = create(SpreadsheetChartAxisArtifactSchema, { minimum: 1, maximum: 100, logBase: 10 });
  c.series[0].xValues = [1, 10, -500, 100]; // Missing pair never enters log domain.
}));
assert.equal(scatterLog.reliability.status, "requires-review");
assert.match(scatterLog.pages[0].svg, /d="M 150 355 L 350 313"/);
const scatterIndependent = paintPpjSceneSvg(scatterFixture(c => {
  c.series.push(create(SpreadsheetChartSeriesArtifactSchema, { name: "Other", xValues: [80, 20], values: [1, 9],
    line: { color: { source: { case: "rgb", value: "00FF00" } }, widthPoints: 2 },
    marker: { symbol: 3, size: 8, fill: { source: { case: "rgb", value: "00FF00" } } } }));
}));
assert.match(scatterIndependent.pages[0].svg, /data-officekit-series="1"[\s\S]*d="M 470 334 L 230 166"/);
const scatterAuto = paintPpjSceneSvg(scatterFixture(c => {
  c.xAxis = undefined; c.series[0].xValues[2] = 1e200;
}));
assert.match(scatterAuto.pages[0].svg, /data-officekit-x-min="0" data-officekit-x-max="100"/);
const scatterOutside = paintPpjSceneSvg(scatterFixture(c => { c.series[0].xValues[0] = -10; }));
assert.match(scatterOutside.pages[0].svg, /data-officekit-x-value="-10" data-officekit-value="0" data-officekit-point-outside-plot="true"/);
assert.doesNotMatch(scatterOutside.pages[0].svg, /circle cx="110"/);
const scatterNone = paintPpjSceneSvg(scatterFixture(c => {
  c.series[0].values.fill(0); c.series[0].missingValueIndexes = [0, 1, 2, 3];
}));
assert.match(scatterNone.pages[0].svg, /No observed data/);
assert.doesNotMatch(scatterNone.pages[0].svg, /data-officekit-point=|data-officekit-line-segment=/);
for (const edit of [
  c => { c.series[0].xValues.pop(); }, c => { c.series[0].bubbleSizes = [1, 2, 3, 4]; },
  c => { c.series[0].missingValueIndexes = [2, 2]; }, c => { c.series[0].values[2] = 9; },
  c => { c.series[0].xValues[0] = NaN; }, c => { c.displayBlanksAs = "span"; },
  c => { c.scatterStyle = "smoothWithMarkers"; }, c => { c.xAxis.minimum = 101; },
  c => { c.xAxis.logBase = 10; }, c => { c.yAxis.logBase = 1; },
  c => { c.series[0].marker.symbol = 9; }, c => { c.scatterStyle = "line"; },
  c => { c.series[0].line = undefined; },
]) {
  const result = paintPpjSceneSvg(scatterFixture(edit));
  assert.equal(result.reliability.status, "failed", JSON.stringify(result.diagnostics));
}

function lineFixture(edit = () => {}) {
  return fixture(scene => {
    const chart = child("line", "chart", { ...frame(100, 100, 500, 300), type: 2,
      categories: ["A", "Missing", "Real zero", "D", "E"],
      yAxis: { minimum: 0, maximum: 5 },
      lineOptions: { grouping: 1, smooth: false },
      series: [{ name: "Observed", values: [2, 0, 0, 4, 5], missingValueIndexes: [1],
        line: { color: { source: { case: "rgb", value: "114477" } }, widthPoints: 2, opacityThousandthPercent: 50000 } }],
    });
    edit(chart.content.value);
    scene.presentation.slides[0].elements = [chart];
  });
}
const lineInput = lineFixture(), lineBytes = toBinary(PresentationPreviewSceneSchema, lineInput.previewScene);
const linePaint = paintPpjSceneSvg(lineInput), lineSvg = linePaint.pages[0].svg;
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, lineInput.previewScene), lineBytes);
assert.match(lineSvg, /data-officekit-chart="line" data-officekit-blank-policy="gap" data-officekit-scale-min="0" data-officekit-scale-max="5"/);
assert.match(lineSvg, /data-officekit-missing-point="1"/);
assert.doesNotMatch(lineSvg, /data-officekit-point="1"|data-officekit-line-segment="0:/);
assert.match(lineSvg, /data-officekit-point="2" data-officekit-value="0"/);
assert.match(lineSvg, /data-officekit-line-segment="2:4" d="M 350 355 L 450 187 L 550 145"[^>]*stroke="#114477" stroke-width="2" stroke-opacity="0.5"/);
assert.match(lineSvg, /data-officekit-review-point="isolated" cx="150" cy="271"/);
assert.ok(linePaint.diagnostics.some(d => d.reason === "preview.scene.paint.chart-isolated-review-point" && d.scenePath.endsWith("series[0].values[0]")));
assert.ok(linePaint.diagnostics.some(d => d.reason === "preview.scene.paint.chart-layout"));
assert.equal(linePaint.reliability.status, "requires-review");
const reversedLine = paintPpjSceneSvg(lineFixture(c => {
  c.xAxis = { ...c.yAxis, minimum: undefined, maximum: undefined, reverse: true };
  c.yAxis.reverse = true;
  c.series[0].line.opacityThousandthPercent = 0;
  c.series[0].marker = create(SpreadsheetChartMarkerArtifactSchema, { symbol: 3, size: 8, fill: { source: { case: "rgb", value: "FF0000" } }, fillOpacityThousandthPercent: 0 });
}));
assert.match(reversedLine.pages[0].svg, /d="M 350 145 L 250 313 L 150 355"[^>]*stroke-opacity="0"/);
assert.match(reversedLine.pages[0].svg, /data-officekit-point="0"[^>]*fill="#FF0000" fill-opacity="0"[^>]*><title>[^<]*<\/title><circle cx="550" cy="229" r="4"/);
const autoRange = paintPpjSceneSvg(lineFixture(c => {
  c.grouping = "none"; // imported native standard, per OpenXmlChartSpaceCodec
  c.yAxis.minimum = undefined; c.yAxis.maximum = undefined;
  c.series[0].values = [10, 0, 20, 0, 30]; c.series[0].missingValueIndexes = [1, 3];
}));
assert.match(autoRange.pages[0].svg, /data-officekit-scale-min="10" data-officekit-scale-max="30"/);
assert.equal((autoRange.pages[0].svg.match(/data-officekit-review-point="isolated"/g) || []).length, 3);
assert.doesNotMatch(autoRange.pages[0].svg, /data-officekit-line-segment=/);
const allMissing = paintPpjSceneSvg(lineFixture(c => {
  c.series[0].values = [0, 0, 0, 0, 0]; c.series[0].missingValueIndexes = [0, 1, 2, 3, 4];
}));
assert.match(allMissing.pages[0].svg, /No observed data/);
assert.doesNotMatch(allMissing.pages[0].svg, /data-officekit-point=|data-officekit-line-segment=/);
const singleton = paintPpjSceneSvg(lineFixture(c => { c.categories = ["only"]; c.series[0].values = [0]; c.series[0].missingValueIndexes = []; }));
assert.match(singleton.pages[0].svg, /data-officekit-review-point="isolated" cx="350" cy="355"/);
const negative = paintPpjSceneSvg(lineFixture(c => {
  c.yAxis.minimum = -5; c.series[0].values = [-5, 0, 0, 4, 5];
}));
assert.match(negative.pages[0].svg, /data-officekit-scale-min="-5"/);
assert.match(negative.pages[0].svg, /data-officekit-point="0" data-officekit-value="-5"/);
const log = paintPpjSceneSvg(lineFixture(c => {
  c.yAxis.minimum = 1; c.yAxis.maximum = 100; c.yAxis.logBase = 10;
  c.series[0].values = [1, 0, 10, 10, 100];
}));
assert.match(log.pages[0].svg, /data-officekit-scale-min="0" data-officekit-scale-max="2" data-officekit-log-base="10"/);
assert.match(log.pages[0].svg, /d="M 350 250 L 450 250 L 550 145"/);
const edgeMissing = paintPpjSceneSvg(lineFixture(c => {
  c.categories = ["lead0", "lead1", "A", "gap", "B", "tail0", "tail1"];
  c.series[0].values = [0, 0, 2, 0, 4, 0, 0];
  c.series[0].missingValueIndexes = [0, 1, 3, 5, 6];
}));
assert.equal((edgeMissing.pages[0].svg.match(/data-officekit-missing-point=/g) || []).length, 5);
assert.deepEqual([...edgeMissing.pages[0].svg.matchAll(/data-officekit-point="(\d+)"/g)].map(m => +m[1]), [2, 4]);
assert.equal((edgeMissing.pages[0].svg.match(/data-officekit-review-point="isolated"/g) || []).length, 2);
assert.doesNotMatch(edgeMissing.pages[0].svg, /data-officekit-line-segment=/);
const multi = paintPpjSceneSvg(lineFixture(c => {
  const second = clone(SpreadsheetChartSeriesArtifactSchema, c.series[0]);
  second.name = "Second & distinct"; second.values = [5, 4, 0, 0, 1]; second.missingValueIndexes = [2];
  second.line.color.source.value = "008800";
  c.series.push(second);
}));
const multiSvg = multi.pages[0].svg;
assert.ok(multiSvg.indexOf('data-officekit-series="0"') < multiSvg.indexOf('data-officekit-series="1"'));
const secondSvg = multiSvg.slice(multiSvg.indexOf('data-officekit-series="1"'));
assert.match(secondSvg, /data-officekit-series-name="Second &amp; distinct"/);
assert.match(secondSvg, /data-officekit-line-segment="0:1" d="M 150 145 L 250 187"[^>]*stroke="#008800"/);
assert.match(secondSvg, /data-officekit-line-segment="3:4" d="M 450 355 L 550 313"/);
assert.match(secondSvg, /data-officekit-missing-point="2"/);
assert.doesNotMatch(secondSvg, /data-officekit-point="2"|data-officekit-line-segment="2:4"/);
assert.match(multiSvg.slice(0, multiSvg.indexOf('data-officekit-series="1"')), /data-officekit-line-segment="2:4"/);
const hiddenAxes = paintPpjSceneSvg(lineFixture(c => {
  c.xAxis = { ...c.yAxis, minimum: undefined, maximum: undefined, visible: false };
  c.yAxis.visible = false;
}));
assert.doesNotMatch(hiddenAxes.pages[0].svg, /data-officekit-axis=|data-officekit-category=|data-officekit-value-tick=/);
assert.match(hiddenAxes.pages[0].svg, /data-officekit-line-segment="2:4"/);
const noAxisLines = paintPpjSceneSvg(lineFixture(c => {
  c.xAxis = { ...c.yAxis, minimum: undefined, maximum: undefined, axisLineVisible: false, tickLabelInterval: 2 };
  c.yAxis.axisLineVisible = false; c.yAxis.tickLabelsVisible = false;
}));
assert.doesNotMatch(noAxisLines.pages[0].svg, /data-officekit-axis=|data-officekit-value-tick=/);
assert.deepEqual([...noAxisLines.pages[0].svg.matchAll(/data-officekit-category="(\d+)"/g)].map(m => +m[1]), [0, 2, 4]);
const outlinedMarker = paintPpjSceneSvg(lineFixture(c => {
  c.series[0].marker = create(SpreadsheetChartMarkerArtifactSchema, { symbol: 3, size: 10,
    fill: { source: { case: "rgb", value: "00FF00" } }, fillOpacityThousandthPercent: 0,
    line: { color: { source: { case: "rgb", value: "112233" } }, widthPoints: 2,
      opacityThousandthPercent: 50000, dashStyle: 1, cap: "round", join: "bevel" } });
}));
assert.match(outlinedMarker.pages[0].svg, /data-officekit-point="0"[^>]*fill="#00FF00" fill-opacity="0" stroke="#112233" stroke-width="2" stroke-opacity="0.5" stroke-linecap="round" stroke-linejoin="bevel"/);
assert.ok(!outlinedMarker.diagnostics.some(d => /marker\.line(?:\.|$)/.test(d.scenePath)));
assert.match(outlinedMarker.pages[0].svg, /data-officekit-line-clip="plot" x="150" y="145" width="400" height="210"/);
assert.match(outlinedMarker.pages[0].svg, /data-officekit-markers="unclipped-at-plot-edge"/);
const invisibleOutline = paintPpjSceneSvg(lineFixture(c => {
  c.series[0].marker = create(SpreadsheetChartMarkerArtifactSchema, { symbol: 3, size: 8,
    fill: { source: { case: "rgb", value: "00FF00" } },
    line: { color: { source: { case: "rgb", value: "112233" } }, widthPoints: 0, opacityThousandthPercent: 0 } });
}));
assert.match(invisibleOutline.pages[0].svg, /data-officekit-point="0"[^>]*stroke="#112233" stroke-width="0" stroke-opacity="0"/);
assert.ok(!invisibleOutline.diagnostics.some(d => /marker\.line(?:\.|$)/.test(d.scenePath)));
const dashedOutline = paintPpjSceneSvg(lineFixture(c => {
  c.series[0].marker = create(SpreadsheetChartMarkerArtifactSchema, { symbol: 3, size: 8,
    line: { color: { source: { case: "rgb", value: "112233" } }, widthPoints: 2, dashStyle: 2 } });
}));
assert.match(dashedOutline.pages[0].svg, /data-officekit-point="0"[^>]*stroke-dasharray="8 6"/);
assert.ok(dashedOutline.diagnostics.some(d => d.reason === "preview.scene.paint.dash-metrics" && d.scenePath.endsWith("series[0].marker.line.dashStyle")));
assert.ok(!dashedOutline.diagnostics.some(d => d.scenePath.endsWith("marker.line.lineStyle")));
const outsideRange = paintPpjSceneSvg(lineFixture(c => {
  c.series[0].values = [6, 0, 0, 4, 5];
  c.series[0].marker = create(SpreadsheetChartMarkerArtifactSchema, { symbol: 3, size: 8 });
}));
assert.match(outsideRange.pages[0].svg, /data-officekit-point="0" data-officekit-value="6" data-officekit-point-outside-plot="true"[^>]*><title>[^<]*<\/title><\/g>/);
assert.match(outsideRange.pages[0].svg, /data-officekit-point="4" data-officekit-value="5"[^>]*><title>[^<]*<\/title><circle cx="550" cy="145" r="4"/);
const outsideNoMarker = paintPpjSceneSvg(lineFixture(c => { c.series[0].values[0] = 6; }));
assert.ok(!outsideNoMarker.diagnostics.some(d => d.reason === "preview.scene.paint.chart-isolated-review-point"));
assert.match(outsideNoMarker.pages[0].svg, /data-officekit-point="0" data-officekit-value="6" data-officekit-point-outside-plot="true"/);
for (const bad of [
  c => { c.series[0].missingValueIndexes = [1, 1]; },
  c => { c.series[0].missingValueIndexes = [3, 1]; },
  c => { c.series[0].missingValueIndexes = [5]; },
  c => { c.series[0].values[1] = 99; },
  c => { c.series[0].values[0] = NaN; },
  c => { c.series[0].values.pop(); },
  c => { c.series[0].xValues = [1, 2, 3, 4, 5]; },
  c => { c.displayBlanksAs = "zero"; },
  c => { c.displayBlanksAs = "span"; },
  c => { c.lineOptions.grouping = 2; },
  c => { c.grouping = "stacked"; },
  c => { c.lineOptions.smooth = true; },
  c => { c.yAxis.logBase = 10; }, // real zero must not be dropped to make log succeed
  c => { c.yAxis.minimum = 5; },
  c => { c.series[0].marker = create(SpreadsheetChartMarkerArtifactSchema, { symbol: 3, size: 0 }); },
]) {
  const failure = paintPpjSceneSvg(lineFixture(bad));
  assert.equal(failure.reliability.status, "failed");
  assert.ok(failure.diagnostics.some(d => d.reason === "preview.scene.paint.chart-semantics"));
  assert.doesNotMatch(failure.pages[0].svg, /data-officekit-line-segment=/);
}
// Both PPJ bar/column lower to native BAR; direction and category/value axes
// must be interpreted from the native contract, not the PPJ type spelling.
function barFixture(edit = () => {}, direction = "column") {
  return lineFixture(c => {
    Object.assign(c, frame(60, 80, 450, 300), { type: 1, barDirection: direction, grouping: "none", gapWidth: 100, overlap: 0, lineOptions: undefined });
    c.categories = ["Positive", "Unknown", "Zero", "Negative"];
    c.yAxis.minimum = -4; c.yAxis.maximum = 8;
    c.series = [create(SpreadsheetChartSeriesArtifactSchema, { name: "First", values: [4, 0, 0, -4], missingValueIndexes: [1],
      seriesFill: { fill: { case: "solidRgb", value: "CC2200" } } }),
    create(SpreadsheetChartSeriesArtifactSchema, { name: "Second", values: [8, 2, 0, -2], missingValueIndexes: [2],
      seriesFill: { fill: { case: "solidRgb", value: "0044CC" } } })];
    edit(c);
  });
}
const barGeometry = svg => [...svg.matchAll(/<rect data-officekit-bar="(\d+)" x="([^"]+)" y="([^"]+)" width="([^"]+)" height="([^"]+)"/g)]
  .map(m => m.slice(1).map(Number));
const columnInput = barFixture(), columnBefore = toBinary(PresentationPreviewSceneSchema, columnInput.previewScene);
const column = paintPpjSceneSvg(columnInput), columnSvg = column.pages[0].svg;
assert.equal(column.reliability.status, "requires-review");
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, columnInput.previewScene), columnBefore);
assert.deepEqual(barGeometry(columnSvg), [[0, 120, 195, 30, 70], [3, 390, 265, 30, 70], [0, 150, 125, 30, 140], [1, 240, 230, 30, 35], [3, 420, 265, 30, 35]]);
assert.equal((columnSvg.match(/data-officekit-missing-point=/g) || []).length, 2);
assert.match(columnSvg, /data-officekit-point="2" data-officekit-value="0" data-officekit-baseline="0"/);
assert.match(columnSvg, /data-officekit-review-point="zero" d="M 300 265 L 330 265"/);
assert.match(columnSvg, /2 missing; 1 zero/);
assert.match(columnSvg, /data-officekit-category="0" x="150" y="347"/);
assert.match(columnSvg, /data-officekit-bar-zero-baseline="review" x1="105" y1="265" x2="465" y2="265"/);
assert.match(columnSvg, /data-officekit-value-tick="0" x="101" y="268"/);
const bars = paintPpjSceneSvg(barFixture(() => {}, "bar"));
assert.deepEqual(barGeometry(bars.pages[0].svg), [[0, 225, 308.75, 120, 17.5], [3, 105, 151.25, 120, 17.5], [0, 225, 291.25, 240, 17.5], [1, 225, 238.75, 60, 17.5], [3, 165, 133.75, 60, 17.5]]);
assert.match(bars.pages[0].svg, /data-officekit-category="0" x="101" y="311.75"/);
for (const direction of ["bar", "column"]) {
  const reversed = paintPpjSceneSvg(barFixture(c => { c.xAxis = { ...c.yAxis, minimum: undefined, maximum: undefined, reverse: true }; c.yAxis.reverse = true; }, direction));
  assert.deepEqual(barGeometry(reversed.pages[0].svg)[0], direction === "column" ? [0, 420, 195, 30, 70] : [0, 225, 133.75, 120, 17.5]);
  const bounds = paintPpjSceneSvg(barFixture(c => { c.yAxis.maximum = 2; }, direction));
  assert.match(bounds.pages[0].svg, /data-officekit-value="4" data-officekit-baseline="0" data-officekit-point-outside-plot="true"/);
  assert.match(bounds.pages[0].svg, /data-officekit-bar-clip="plot"[^>]*overflow="hidden"/);
  const empty = paintPpjSceneSvg(barFixture(c => { for (const entry of c.series) { entry.values = [0, 0, 0, 0]; entry.missingValueIndexes = [0, 1, 2, 3]; } }, direction));
  assert.equal(empty.reliability.status, "requires-review");
  assert.match(empty.pages[0].svg, /No observed data; 8 missing; 0 zero/);
  assert.doesNotMatch(empty.pages[0].svg, /data-officekit-bar=|data-officekit-point=|data-officekit-review-point=/);
  const hiddenAxes = paintPpjSceneSvg(barFixture(c => { c.showCategoryAxis = false; c.showValueAxis = false; }, direction));
  assert.doesNotMatch(hiddenAxes.pages[0].svg, /data-officekit-axis=|data-officekit-category=|data-officekit-value-tick=|data-officekit-bar-zero-baseline=/);
  for (const overlap of [-100, 0, 50, 100]) for (const gap of [0, 100, 500]) {
    const result = paintPpjSceneSvg(barFixture(c => { c.overlap = overlap; c.gapWidth = gap; }, direction));
    const [first, , second] = barGeometry(result.pages[0].svg);
    const width = direction === "column" ? first[3] : first[4], band = direction === "column" ? 90 : 52.5;
    assert.ok(Math.abs(width - band / (2 - overlap / 100 + gap / 100)) < 1e-10);
    assert.ok(Math.abs(Math.abs(first[direction === "column" ? 1 : 2] - second[direction === "column" ? 1 : 2]) - width * (1 - overlap / 100)) < 1e-10);
  }
}
const positiveBars = paintPpjSceneSvg(barFixture(c => { c.yAxis = undefined; c.series[0].values = [4, 0, 3, 2]; c.series[1].values = [8, 2, 0, 1]; }));
assert.match(positiveBars.pages[0].svg, /data-officekit-scale-min="0" data-officekit-scale-max="8"/);
const defaultBars = paintPpjSceneSvg(barFixture(c => { c.barDirection = ""; c.gapWidth = undefined; c.overlap = undefined; }));
assert.match(defaultBars.pages[0].svg, /data-officekit-chart="column"[^>]*data-officekit-gap-width="150" data-officekit-overlap="0"/);
const pointBars = paintPpjSceneSvg(barFixture(c => {
  c.series[0].pointStyles = [create(SpreadsheetChartPointStyleArtifactSchema, { index: 0, fill: { fill: { case: "solidRgb", value: "00AA44" }, opacityThousandthPercent: 0 },
    line: { color: { source: { case: "rgb", value: "223344" } }, widthPoints: 0, opacityThousandthPercent: 0 } }),
  create(SpreadsheetChartPointStyleArtifactSchema, { index: 3, fill: { fill: { case: "noFill", value: true } } })];
}));
assert.match(pointBars.pages[0].svg, /data-officekit-bar="0"[^>]*fill="#00AA44" fill-opacity="0" stroke="#223344" stroke-width="0" stroke-opacity="0"/);
assert.match(pointBars.pages[0].svg, /data-officekit-bar="3"[^>]*fill="none"/);
const residualBar = paintPpjSceneSvg(barFixture(c => { c.hasLegend = true; c.series[0].valuesFormatCode = "0%"; }));
assert.ok(residualBar.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith("valuesFormatCode")));
for (const bad of [
  c => { c.barDirection = "diagonal"; }, c => { c.grouping = "unknown"; }, c => { c.grouping = "percent-stacked"; },
  c => { c.gapWidth = 501; }, c => { c.overlap = -101; }, c => { c.overlap = 101; },
  c => { c.yAxis.logBase = 10; }, c => { c.yAxis.minimum = 8; }, c => { c.secondaryYAxis = c.yAxis; },
  c => { c.series[0].values[1] = 2; }, c => { c.series[0].missingValueIndexes = [2, 1]; },
  c => { c.series[0].values[0] = NaN; }, c => { c.series[0].values.pop(); },
  c => { c.series[0].xValues = [1, 2, 3, 4]; }, c => { c.displayBlanksAs = "zero"; }, c => { c.displayBlanksAs = "span"; },
  c => { c.series[0].pointStyles = [create(SpreadsheetChartPointStyleArtifactSchema, { index: 1 })]; },
]) {
  const failure = paintPpjSceneSvg(barFixture(bad));
  assert.equal(failure.reliability.status, "failed");
  assert.ok(failure.diagnostics.some(d => d.reason === "preview.scene.paint.chart-semantics"));
  assert.doesNotMatch(failure.pages[0].svg, /data-officekit-bar=|data-officekit-review-point="zero"/);
}
// Stacking changes value coordinates, independently of declared bar overlap.
for (const direction of ["column", "bar"]) {
  const stackedFixture = (edit = () => {}) => barFixture(c => {
    c.grouping = "stacked"; c.overlap = 100; c.yAxis = undefined; edit(c);
  }, direction);
  const input = stackedFixture(), before = toBinary(PresentationPreviewSceneSchema, input.previewScene);
  const result = paintPpjSceneSvg(input), svg = result.pages[0].svg;
  assert.equal(result.reliability.status, "requires-review");
  assert.deepEqual(toBinary(PresentationPreviewSceneSchema, input.previewScene), before);
  assert.match(svg, /data-officekit-scale-min="-6" data-officekit-scale-max="12" data-officekit-grouping="stacked"/);
  assert.match(svg, /data-officekit-value="8" data-officekit-baseline="4" data-officekit-stack-end="12"/);
  assert.match(svg, /data-officekit-value="-2" data-officekit-baseline="-4" data-officekit-stack-end="-6"/);
  assert.match(svg, /data-officekit-point="1" data-officekit-value="2" data-officekit-stack-position="unknown"/);
  assert.match(svg, /data-officekit-point="2" data-officekit-value="0" data-officekit-stack-position="unknown"/);
  assert.equal((svg.match(/data-officekit-incomplete-stack=/g) || []).length, 2);
  assert.equal((svg.match(/data-officekit-missing-point=/g) || []).length, 2);
  assert.doesNotMatch(svg, /data-officekit-bar="[12]"|data-officekit-review-point="zero"/);
  const geometry = barGeometry(svg);
  const expected = direction === "column"
    ? [[0, 127.5, 125 + 8 * 210 / 18, 45, 4 * 210 / 18], [3, 397.5, 265, 45, 4 * 210 / 18], [0, 127.5, 125, 45, 8 * 210 / 18], [3, 397.5, 265 + 4 * 210 / 18, 45, 2 * 210 / 18]]
    : [[0, 225, 295.625, 80, 26.25], [3, 145, 138.125, 80, 26.25], [0, 305, 295.625, 160, 26.25], [3, 105, 138.125, 40, 26.25]];
  assert.equal(geometry.length, expected.length);
  geometry.forEach((row, i) => row.forEach((v, j) => assert.ok(Math.abs(v - expected[i][j]) < 1e-9)));
  const reversed = paintPpjSceneSvg(stackedFixture(c => { c.xAxis = create(SpreadsheetChartAxisArtifactSchema, { reverse: true }); c.yAxis = create(SpreadsheetChartAxisArtifactSchema, { reverse: true }); }));
  barGeometry(reversed.pages[0].svg).forEach((row, i) => {
    const [, x, y, w, h] = geometry[i];
    row.forEach((v, j) => assert.ok(Math.abs(v - [geometry[i][0], 570 - x - w, 460 - y - h, w, h][j]) < 1e-9));
  });
  const signed = paintPpjSceneSvg(stackedFixture(c => {
    c.categories = ["Mixed", "Zero"];
    c.series[0].values = [4, 2]; c.series[0].missingValueIndexes = [];
    c.series[1].values = [-2, 0]; c.series[1].missingValueIndexes = [];
    c.series.push(create(SpreadsheetChartSeriesArtifactSchema, { name: "Third", values: [3, -1] }));
  }));
  assert.match(signed.pages[0].svg, /data-officekit-value="3" data-officekit-baseline="4" data-officekit-stack-end="7"/);
  assert.match(signed.pages[0].svg, /data-officekit-value="-2" data-officekit-baseline="0" data-officekit-stack-end="-2"/);
  assert.match(signed.pages[0].svg, /data-officekit-value="0" data-officekit-baseline="2" data-officekit-stack-end="2"/);
  assert.match(signed.pages[0].svg, /data-officekit-review-point="zero"/);
  for (const overlap of [-100, 0, 100]) {
    const shifted = paintPpjSceneSvg(stackedFixture(c => { c.overlap = overlap; }));
    assert.match(shifted.pages[0].svg, /data-officekit-value="8" data-officekit-baseline="4" data-officekit-stack-end="12"/);
    const bars = barGeometry(shifted.pages[0].svg), extentIndex = direction === "column" ? 3 : 4;
    assert.ok(Math.abs(bars[0][extentIndex] - (direction === "column" ? 90 : 52.5) / (3 - overlap / 100)) < 1e-9);
  }
  for (const values of [[1e308, 1e308], [1e20, 1]]) {
    const bad = paintPpjSceneSvg(stackedFixture(c => { c.series[0].values[0] = values[0]; c.series[1].values[0] = values[1]; }));
    assert.equal(bad.reliability.status, "failed");
    assert.doesNotMatch(bad.pages[0].svg, /data-officekit-bar=/);
  }
}
for (const direction of ["column", "bar"]) {
  const percentFixture = (edit = () => {}) => barFixture(c => {
    c.grouping = "percent-stacked"; c.overlap = 100; c.yAxis = undefined;
    c.categories[3] = "Other";
    c.series[0].values = [1, 0, 0, 3];
    c.series[1].values = [3, 2, 0, 1]; c.series[1].missingValueIndexes = [];
    edit(c);
  }, direction);
  const input = percentFixture(), before = toBinary(PresentationPreviewSceneSchema, input.previewScene);
  const result = paintPpjSceneSvg(input), svg = result.pages[0].svg;
  assert.equal(result.reliability.status, "requires-review");
  assert.deepEqual(toBinary(PresentationPreviewSceneSchema, input.previewScene), before);
  assert.match(svg, /data-officekit-scale-min="0" data-officekit-scale-max="1" data-officekit-grouping="percent-stacked"/);
  assert.match(svg, /data-officekit-value="1" data-officekit-baseline="0" data-officekit-stack-end="0.25" data-officekit-fraction="0.25"/);
  assert.match(svg, /data-officekit-value="3" data-officekit-baseline="0.25" data-officekit-stack-end="1" data-officekit-fraction="0.75"/);
  assert.match(svg, /data-officekit-value-tick="0.5"[^>]*>50%<\/text>/);
  assert.match(svg, /data-officekit-incomplete-stack="1"/);
  assert.match(svg, /data-officekit-zero-total-stack="2"/);
  assert.equal((svg.match(/data-officekit-stack-position="zero-total"/g) || []).length, 2);
  assert.doesNotMatch(svg, /data-officekit-bar="[12]"|data-officekit-review-point="zero"/);
  const expected = direction === "column"
    ? [[0, 127.5, 282.5, 45, 52.5], [3, 397.5, 177.5, 45, 157.5], [0, 127.5, 125, 45, 157.5], [3, 397.5, 125, 45, 52.5]]
    : [[0, 105, 295.625, 90, 26.25], [3, 105, 138.125, 270, 26.25], [0, 195, 295.625, 270, 26.25], [3, 375, 138.125, 90, 26.25]];
  assert.deepEqual(barGeometry(svg), expected);
  const reversed = paintPpjSceneSvg(percentFixture(c => { c.xAxis = create(SpreadsheetChartAxisArtifactSchema, { reverse: true }); c.yAxis = create(SpreadsheetChartAxisArtifactSchema, { reverse: true }); }));
  assert.deepEqual(barGeometry(reversed.pages[0].svg), expected.map(([i, x, y, w, h]) => [i, 570 - x - w, 460 - y - h, w, h]));
  const big = paintPpjSceneSvg(percentFixture(c => { c.series[0].values[0] = 1e308; c.series[1].values[0] = 1e308; }));
  assert.equal(big.reliability.status, "requires-review");
  assert.match(big.pages[0].svg, /data-officekit-fraction="0.5"/);
  const zero = paintPpjSceneSvg(percentFixture(c => { c.series[0].values[0] = 0; }));
  assert.match(zero.pages[0].svg, /data-officekit-value="0" data-officekit-baseline="0" data-officekit-stack-end="0" data-officekit-fraction="0"/);
  assert.match(zero.pages[0].svg, /data-officekit-review-point="zero"/);
  const bounded = paintPpjSceneSvg(percentFixture(c => { c.yAxis = create(SpreadsheetChartAxisArtifactSchema, { minimum: 0, maximum: .5 }); }));
  assert.match(bounded.pages[0].svg, /data-officekit-point-outside-plot="true" data-officekit-stack-end="1" data-officekit-fraction="0.75"/);
  for (const invalid of [c => { c.series[0].values[0] = -1; }, c => { c.series[0].values[0] = 1e-308; c.series[1].values[0] = 1e308; },
    c => { c.series[0].values[0] = 1e20; c.series[1].values[0] = 1; }, c => { c.displayBlanksAs = "zero"; }]) {
    const failure = paintPpjSceneSvg(percentFixture(invalid));
    assert.equal(failure.reliability.status, "failed"); assert.doesNotMatch(failure.pages[0].svg, /data-officekit-bar=/);
  }
}
function circularFixture(edit = () => {}, type = 3) {
  return lineFixture(c => {
    c.type = type; c.yAxis = undefined; c.lineOptions = undefined;
    c.categories = ["Small", "Unknown", "Real zero", "Large"];
    c.series[0].values = [1, 0, 0, 9]; c.series[0].missingValueIndexes = [1]; c.series[0].line = undefined;
    c.series[0].pointStyles = [create(SpreadsheetChartPointStyleArtifactSchema, { index: 0,
      fill: { fill: { case: "solidRgb", value: "CC2200" } } }), create(SpreadsheetChartPointStyleArtifactSchema, { index: 3,
      fill: { fill: { case: "solidRgb", value: "0044CC" } } })];
    edit(c);
  });
}
const sliceAttributes = svg => [...svg.matchAll(/data-officekit-point="(\d+)" data-officekit-value="([^"]+)" data-officekit-fraction="([^"]+)" data-officekit-start-angle="([^"]+)" data-officekit-sweep-angle="([^"]+)"/g)]
  .map(m => ({ index: +m[1], value: +m[2], fraction: +m[3], start: +m[4], sweep: +m[5] }));
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-10, `${actual} != ${expected}`);
const pieReceipt = circularFixture(), pieBefore = toBinary(PresentationPreviewSceneSchema, pieReceipt.previewScene);
const pie = paintPpjSceneSvg(pieReceipt), pieSvg = pie.pages[0].svg;
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, pieReceipt.previewScene), pieBefore);
assert.match(pieSvg, /data-officekit-chart="pie" data-officekit-first-slice-angle="0" data-officekit-hole-size="0"/);
const piePoints = sliceAttributes(pieSvg);
assert.deepEqual(piePoints.map(p => [p.index, p.value]), [[0, 1], [2, 0], [3, 9]]);
near(piePoints[0].fraction, .1); near(piePoints[0].sweep, 36); near(piePoints[2].sweep, 324);
assert.match(pieSvg, /data-officekit-slice="0"[^>]*fill="#CC2200"/);
assert.match(pieSvg, /data-officekit-slice="3"[^>]*fill="#0044CC"/);
assert.doesNotMatch(pieSvg, /data-officekit-slice="[12]"|data-officekit-point="1"/);
assert.match(pieSvg, /data-officekit-missing-point="1"/);
assert.match(pieSvg, /Observed values only; 1 missing; 1 zero/);
assert.ok(pie.diagnostics.some(d => d.reason === "preview.scene.paint.chart-missing-share"));
assert.equal(pie.reliability.status, "requires-review");
const swappedPie = paintPpjSceneSvg(circularFixture(c => { c.series[0].values = [9, 0, 0, 1]; }));
near(sliceAttributes(swappedPie.pages[0].svg)[0].sweep, 324);
assert.notEqual(pieSvg.match(/data-officekit-slice="0" d="([^"]+)"/)[1], swappedPie.pages[0].svg.match(/data-officekit-slice="0" d="([^"]+)"/)[1]);
const doughnut = paintPpjSceneSvg(circularFixture(c => { c.firstSliceAngle = 90; c.doughnutHoleSize = 60; }, 5));
assert.match(doughnut.pages[0].svg, /data-officekit-chart="doughnut" data-officekit-first-slice-angle="90" data-officekit-hole-size="60"/);
near(sliceAttributes(doughnut.pages[0].svg)[0].start, 90);
assert.equal((doughnut.pages[0].svg.match(/ A /g) || []).length, 8);
const defaultHole = paintPpjSceneSvg(circularFixture(() => {}, 5));
assert.match(defaultHole.pages[0].svg, /data-officekit-hole-size="50"/);
for (const type of [3, 5]) {
  const full = paintPpjSceneSvg(circularFixture(c => {
    c.categories = ["All"]; c.series[0].values = [4]; c.series[0].missingValueIndexes = []; c.series[0].pointStyles = [];
  }, type));
  near(sliceAttributes(full.pages[0].svg)[0].sweep, 360);
  assert.equal((full.pages[0].svg.match(/ A /g) || []).length, type === 3 ? 2 : 4);
  const noPositive = paintPpjSceneSvg(circularFixture(c => { c.series[0].values = [0, 0, 0, 0]; }, type));
  assert.match(noPositive.pages[0].svg, /No positive observed data/);
  assert.match(noPositive.pages[0].svg, /1 missing; 3 zero/);
  assert.doesNotMatch(noPositive.pages[0].svg, /data-officekit-slice=/);
  const noObservations = paintPpjSceneSvg(circularFixture(c => {
    c.series[0].values = [0, 0, 0, 0]; c.series[0].missingValueIndexes = [0, 1, 2, 3];
    c.series[0].pointStyles = [];
  }, type));
  assert.equal(noObservations.reliability.status, "requires-review");
  assert.doesNotMatch(noObservations.pages[0].svg, /data-officekit-point=|data-officekit-slice=/);
  assert.match(noObservations.pages[0].svg, /4 missing; 0 zero/);
  const invalidMissingStyles = paintPpjSceneSvg(circularFixture(c => {
    c.series[0].values = [0, 0, 0, 0]; c.series[0].missingValueIndexes = [0, 1, 2, 3];
  }, type));
  assert.equal(invalidMissingStyles.reliability.status, "failed");
  assert.ok(invalidMissingStyles.diagnostics.some(d => d.reason === "preview.scene.paint.chart-semantics" && d.scenePath.endsWith("pointStyles[0].index")));
  assert.doesNotMatch(invalidMissingStyles.pages[0].svg, /data-officekit-slice=/);
}
const circularAlpha = paintPpjSceneSvg(circularFixture(c => {
  c.series[0].pointStyles[0].fill.opacityThousandthPercent = 0;
  c.series[0].pointStyles[1].fill = create(SpreadsheetChartSurfaceFillSchema, { fill: { case: "noFill", value: true } });
}));
assert.match(circularAlpha.pages[0].svg, /data-officekit-slice="0"[^>]*fill="#CC2200" fill-opacity="0"/);
assert.match(circularAlpha.pages[0].svg, /data-officekit-slice="3"[^>]*fill="none"/);
const largeCircular = paintPpjSceneSvg(circularFixture(c => { c.series[0].values = [1e308, 0, 0, 1e308]; }));
near(sliceAttributes(largeCircular.pages[0].svg)[0].fraction, .5);
const explosionPoints = svg => [...svg.matchAll(/data-officekit-point="(\d+)"[^>]*data-officekit-explosion="([^"]+)" data-officekit-explosion-owner="([^"]+)" data-officekit-offset-x="([^"]+)" data-officekit-offset-y="([^"]+)"/g)]
  .map(m => ({ index: +m[1], value: +m[2], owner: m[3], x: +m[4], y: +m[5] }));
for (const type of [3, 5]) {
  const makeExploded = (pointZero = false, angle = 0, amount = 100) => circularFixture(c => {
    c.series[0].values = [1, 0, 0, 1]; c.series[0].explosion = amount;
    c.firstSliceAngle = angle;
    if (type === 5) c.doughnutHoleSize = 60;
    if (pointZero) c.series[0].pointStyles[0].explosion = 0;
  }, type);
  for (const amount of [0, 25, 100, 400]) {
    const receipt = makeExploded(false, 0, amount), before = toBinary(PresentationPreviewSceneSchema, receipt.previewScene);
    const result = paintPpjSceneSvg(receipt), svg = result.pages[0].svg;
    assert.equal(result.reliability.status, "requires-review");
    assert.deepEqual(toBinary(PresentationPreviewSceneSchema, receipt.previewScene), before);
    const points = explosionPoints(svg), thickness = type === 5 ? .4 : 1;
    const radius = 102 / (1 + thickness * amount / 100), offset = radius * thickness * amount / 100;
    near(+svg.match(/data-officekit-outer-radius="([^"]+)"/)[1], radius);
    near(points[0].x, offset); near(points[0].y, 0);
    near(points[2].x, -offset); near(points[2].y, 0);
    near(points[1].x, 0); near(points[1].y, 0); // Real zero cannot move or resize the plot.
    assert.deepEqual(points.map(p => p.owner), ["series", "series", "series"]);
    assert.deepEqual(sliceAttributes(svg).map(p => p.fraction), [.5, 0, .5]);
    assert.equal(result.diagnostics.some(d => d.reason === "preview.scene.paint.chart-explosion-layout"), amount > 0);
    assert.equal((svg.match(/transform="translate\([^)]*\)"/g) || []).length, amount > 0 ? 2 : 0);
    assert.ok(radius + offset <= 102 + 1e-10, "whole exploded circle remains within review plot");
  }
  const overridden = paintPpjSceneSvg(makeExploded(true)), points = explosionPoints(overridden.pages[0].svg);
  assert.equal(points[0].value, 0); assert.equal(points[0].owner, "point"); near(points[0].x, 0); near(points[0].y, 0);
  assert.equal(points[2].value, 100); assert.equal(points[2].owner, "series"); assert.ok(points[2].x < 0);
  const rotated = explosionPoints(paintPpjSceneSvg(makeExploded(false, 90)).pages[0].svg);
  near(rotated[0].x, 0); assert.ok(rotated[0].y > 0); near(rotated[2].x, 0); assert.ok(rotated[2].y < 0);
  const zeroOnly = paintPpjSceneSvg(circularFixture(c => {
    c.series[0].pointStyles.push(create(SpreadsheetChartPointStyleArtifactSchema, { index: 2, explosion: 400 }));
    c.series[0].pointStyles.sort((a, b) => a.index - b.index);
  }, type));
  near(+zeroOnly.pages[0].svg.match(/data-officekit-outer-radius="([^"]+)"/)[1], 102);
  assert.ok(!zeroOnly.diagnostics.some(d => d.reason === "preview.scene.paint.chart-explosion-layout"));
}
for (const edit of [
  c => { c.series[0].values[0] = -1; }, c => { c.series[0].values[0] = NaN; },
  c => { c.series[0].values.pop(); }, c => { c.series[0].missingValueIndexes = [1, 1]; },
  c => { c.series[0].values[1] = 2; }, c => { c.firstSliceAngle = 361; },
  c => { c.doughnutHoleSize = 50; }, c => { c.series[0].explosion = 401; },
  c => { c.series[0].pointStyles[0].explosion = 401; },
  c => { c.series[0].pointStyles[1].index = 0; }, c => { c.series[0].pointStyles[1].index = 9; },
  c => { c.series[0].pointStyles[0].index = 1; },
  c => { c.series[0].pointStyles.reverse(); },
  c => { c.series.push(clone(SpreadsheetChartSeriesArtifactSchema, c.series[0])); },
  c => { c.displayBlanksAs = "zero"; }, c => { c.displayBlanksAs = "span"; },
  c => { c.series[0].values = [1e-308, 0, 0, 1e308]; },
]) {
  const failed = paintPpjSceneSvg(circularFixture(edit));
  assert.equal(failed.reliability.status, "failed");
  assert.ok(failed.diagnostics.some(d => d.reason === "preview.scene.paint.chart-semantics"));
  assert.doesNotMatch(failed.pages[0].svg, /data-officekit-slice=/);
}
for (const hole of [0, 9, 91, 100]) {
  const failed = paintPpjSceneSvg(circularFixture(c => { c.doughnutHoleSize = hole; }, 5));
  assert.equal(failed.reliability.status, "failed");
}
const limits = paintPpjSceneSvg(fixture(scene => {
  scene.presentation.slides[0].elements[0].content.value.shadow = create(
    PresentationElementSchema.fields.find(f => f.localName === "shape").message.fields.find(f => f.localName === "shadow").message,
    { distanceEmu: 0n, colorRgb: "000000" });
  scene.presentation.slides[0].elements[0].content.value.futurePaint = 0;
  scene.presentation.slides[0].elements[2].content.value.customPaths[0].commands[1].command.value.xReference = "guide";
  scene.presentation.slides[0].elements[4].content.value.childWidthEmu = 0n;
}));
assert.equal(limits.reliability.status, "failed");
assert.ok(limits.diagnostics.some(d => d.scenePath.endsWith(".shape.shadow")));
assert.ok(limits.diagnostics.some(d => d.reason === "preview.scene.unknown-field" && d.scenePath.endsWith(".futurePaint")));
assert.ok(limits.diagnostics.some(d => d.reason === "preview.scene.paint.path"));
assert.ok(limits.diagnostics.some(d => d.reason === "preview.scene.paint.failed"));
assert.match(limits.pages[0].svg, /path unavailable/);
assert.match(limits.pages[0].svg, /native drawing unavailable/);
assert.doesNotMatch(limits.pages[0].svg, /data-officekit-path="0"/);
assert.match(limits.pages[0].svg, /data-officekit-review="failed"/);

const registry = JSON.parse(readFileSync(new URL("../src/ppj/capability-registry.json", import.meta.url)));
assert.equal(registry.previewScene.painter, "src/ppj/preview-scene-svg.mjs");
assert.equal(registry.previewScene.painterTest, "test/ppj-preview-scene-svg.mjs");
// A new process proves an actual in-memory draw is independent of codec,
// filesystem asset reads, legacy facade and raster/specialist loading.
const lazy = spawnSync(process.execPath, ["--input-type=module", "-e", `
import assert from "node:assert/strict";
import { registerHooks } from "node:module";
const hooks=registerHooks({resolve(s,c,next){assert.doesNotMatch(s,/^(sharp|canvas|jszip|mupdf|playwright)$|office-kit-native-client|codecs\\/office-kit-runtime|\\/presentation\\//);return next(s,c);}});
const { fromBinary }=await import("@bufbuild/protobuf");
const { PresentationPreviewSceneSchema }=await import("./src/generated/office_kit/artifact/v1/office_artifact_pb.js");
const { paintPpjSceneSvg }=await import("./src/ppj/preview-scene-svg.mjs");
const r=JSON.parse(Buffer.from(process.env.SCENE_FIXTURE,"base64"));
for(const k of ["file","programJson"]) r[k]=Buffer.from(r[k],"base64");
r.previewScene=fromBinary(PresentationPreviewSceneSchema,Buffer.from(r.previewScene,"base64"));
for(const a of r.assets)a.data=Buffer.from(a.data,"base64");
assert.match(paintPpjSceneSvg(r).pages[0].svg,/<ellipse/); hooks.deregister();
`], { encoding: "utf8", env: { ...process.env, SCENE_FIXTURE: Buffer.from(JSON.stringify({ ...receipt,
  file: receipt.file.toString("base64"), programJson: receipt.programJson.toString("base64"), previewScene: Buffer.from(before).toString("base64"),
  assets: receipt.assets.map(a => ({ ...a, data: a.data.toString("base64") })),
})).toString("base64") } });
assert.equal(lazy.status, 0, lazy.stderr || lazy.stdout);
console.log("ppj native scene SVG foundations ok: shapes/text/images/groups, connectors/tables, native line and circular proportions, explicit residuals and leaf drawing");
