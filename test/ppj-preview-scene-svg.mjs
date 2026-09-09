import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { create, clone, toBinary } from "@bufbuild/protobuf";
import { PresentationPreviewSceneSchema, PresentationElementSchema,
  PresentationCustomGeometryPathSchema } from "../src/generated/office_kit/artifact/v1/office_artifact_pb.js";
import { nativePathData, paintPpjSceneSvg } from "../src/ppj/preview-scene-svg.mjs";

const sha = data => createHash("sha256").update(data).digest("hex");
const emu = value => BigInt(Math.round(value * 12700));
const frame = (x, y, width, height) => ({ leftEmu: emu(x), topEmu: emu(y), widthEmu: emu(width), heightEmu: emu(height) });
const child = (id, kind, value, hidden) => create(PresentationElementSchema, { id, hidden, content: { case: kind, value } });
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
for (const bad of [
  path({ commands: [command("lineTo", { x: 0n, y: 0n })] }),
  path({ commands: [command("moveTo", { xReference: "guide" })] }),
  path({ commands: [command("moveTo", { x: 2n ** 63n })] }),
  path({ commands: [command("moveTo", {}), command("arcTo", {})] }),
]) assert.throws(() => nativePathData(bad, { x: 0, y: 0, width: 100, height: 100 }));

// Synthetic native receipts isolate drawing semantics; real compiler coverage
// lives in the explicitly selected NativeAOT integration test.
function fixture(edit = () => {}) {
  const file = Buffer.from("candidate"), programJson = Buffer.from('{"pages":[]}');
  const data = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLttAAAAABJRU5ErkJggg==", "base64");
  const nodes = [
    child("box", "shape", { ...frame(10, 40, 200, 100), geometry: "textbox", fillRgb: "114477", lineRgb: "AA2200", lineWidthEmu: emu(2),
      fillOpacityThousandthPercent: 0, text: "WRONG FLATTENED TEXT", textBody: { paragraphs: [{ alignment: "center",
        defaultRunStyle: { case: "defaultRunProperties", value: { fontSizePoints: 16, bold: true, color: { case: "colorRgb", value: "123456" } } },
        runs: [{ content: { case: "text", value: "A & " }, bold: false }, { content: { case: "text", value: "B" }, fontSizePoints: 12.5 },
          { content: { case: "lineBreak", value: true } }, { content: { case: "text", value: "C" }, colorOpacityThousandthPercent: 0 }],
      }] } }),
    child("diamond", "shape", { ...frame(220, 40, 100, 100), geometry: "flowChartDecision", fillRgb: "33AA44", text: "Decision" }),
    child("vector", "shape", { ...frame(10, 150, 200, 100), geometry: "custom", fillRgb: "CC5500", customPaths: [clone(PresentationCustomGeometryPathSchema, polygon)] }),
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
    });
  }
  bind(scene.presentation.slides[0].elements, "$.presentation.slides[0].elements");
  scene.sha256 = sha(toBinary(PresentationPreviewSceneSchema, { ...scene, sha256: "" }));
  return { file, programJson, programSha256: sha(programJson), outputSha256: sha(file), sourceBound: scene.origin === 2, previewScene: scene,
    assets: [{ id: "public-asset", mimeType: "image/png", sha256: sha(data), data }] };
}
const receipt = fixture(), before = toBinary(PresentationPreviewSceneSchema, receipt.previewScene);
const painted = paintPpjSceneSvg(receipt), svg = painted.pages[0].svg;
assert.deepEqual(toBinary(PresentationPreviewSceneSchema, receipt.previewScene), before);
assert.equal(painted.scene, receipt.previewScene);
assert.match(svg, /<rect x="10" y="40" width="200" height="100" fill="#114477" fill-opacity="0"/);
assert.match(svg, /stroke-width="2"/);
assert.match(svg, /<text x="110"[^>]*text-anchor="middle"[^>]*><tspan[^>]*font-weight="normal"[^>]*>A &amp; <\/tspan><tspan[^>]*font-size="12.5"[^>]*font-weight="bold"[^>]*>B<\/tspan><\/text>/);
assert.doesNotMatch(svg, /WRONG FLATTENED|HIDDEN TEXT/);
assert.match(svg, /d="M 270 40 L 320 90 L 270 140 L 220 90 Z"/);
assert.match(svg, /data-officekit-path="0" d="M 10 150 L 210 150 C 210 170 50 250 10 250 Q 110 200 10 150 Z" stroke="none"/);
assert.match(svg, /translate\(355 90\) rotate\(-90\) scale\(-1 1\) translate\(-355 -90\)/);
assert.match(svg, /<image x="330" y="40" width="50" height="100"[^>]*preserveAspectRatio="none" opacity="0"/);
assert.equal(Buffer.from(svg.match(/href="data:image\/png;base64,([^"]+)"/)[1], "base64").compare(receipt.assets[0].data), 0);
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
  c.series[0].marker = { symbol: 3, size: 8, fill: { source: { case: "rgb", value: "FF0000" } }, fillOpacityThousandthPercent: 0 };
}));
assert.match(reversedLine.pages[0].svg, /d="M 350 145 L 250 313 L 150 355"[^>]*stroke-opacity="0"/);
assert.match(reversedLine.pages[0].svg, /data-officekit-point="0"[^>]*fill="#FF0000" fill-opacity="0"[^>]*><title>[^<]*<\/title><circle cx="550" cy="229" r="4"/);
const autoRange = paintPpjSceneSvg(lineFixture(c => {
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
  c => { c.lineOptions.smooth = true; },
  c => { c.yAxis.logBase = 10; }, // real zero must not be dropped to make log succeed
  c => { c.yAxis.minimum = 5; },
]) {
  const failure = paintPpjSceneSvg(lineFixture(bad));
  assert.equal(failure.reliability.status, "failed");
  assert.ok(failure.diagnostics.some(d => d.reason === "preview.scene.paint.chart-semantics"));
  assert.doesNotMatch(failure.pages[0].svg, /data-officekit-line-segment=/);
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
console.log("ppj native scene SVG foundations ok: shapes/text/images/groups, directed connectors, unequal/merged tables, explicit residuals and leaf drawing");
