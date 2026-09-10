import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { assessPpjPreviewInput } from "../src/ppj/preview-input-assessment.mjs";
import { paintPpjSceneSvg } from "../src/ppj/preview-scene-svg.mjs";
import { previewSceneFixture, nativeElement, emuFrame } from "./helpers/ppj-preview-scene-fixture.mjs";

const frame = { x: 1, y: 2, width: 100, height: 60 };
const text = { type: "text", id: "text", frame, text: "Text" };
const deck = (elements) => ({ pages: [{ id: "p1", role: "test", elements }] });
const reasonsAt = (result, path) => result.diagnostics.filter((diagnostic) => diagnostic.path === path).map((diagnostic) => diagnostic.reason);
const sample = deck([{ type: "group", id: "group", frame, elements: [text] }]);
const original = JSON.stringify(sample);
const result = assessPpjPreviewInput(sample);
assert.equal(JSON.stringify(sample), original);
assert.equal(result.children[0].children[0].children[0].id, "text");
assert.ok(reasonsAt(result, "$.pages[0].elements[0].elements[0].text").includes("preview.text.unassessed"));
assert.ok(!result.diagnostics.some((d) => d.path.endsWith(".frame.x")));

const sizedConnector = assessPpjPreviewInput(deck([{ type: "connector", id: "edge", frame,
  connectorType: "straight", from: { x: 1, y: 2 }, to: { x: 101, y: 62 },
  stroke: { color: "#112233", width: 2 }, startArrow: "open", endArrow: "triangle",
  startArrowWidth: "sm", startArrowLength: "lg", endArrowWidth: "med", endArrowLength: "sm" }]));
assert.notEqual(sizedConnector.status, "supported");
for (const field of ["startArrowWidth", "startArrowLength", "endArrowWidth", "endArrowLength"])
  assert.ok(sizedConnector.diagnostics.some(d => d.path.endsWith(`.${field}`) && d.status !== "supported"));

const siteConnector = assessPpjPreviewInput(deck([{ type: "connector", id: "site-edge", frame,
  connectorType: "straight", from: { element: "target", connectionSite: 0 }, to: { x: 101, y: 62 },
  stroke: { color: "#112233", width: 2 } }]));
assert.ok(siteConnector.diagnostics.some(d => d.path.endsWith(".from.connectionSite") && d.status !== "supported"));

const uprightText = assessPpjPreviewInput(deck([{ ...text, style: { upright: false } }]));
assert.ok(uprightText.diagnostics.some(d => d.path.endsWith(".style.upright") && d.status !== "supported"));
const columnText = assessPpjPreviewInput(deck([{ ...text, style: { columnDirection: "left-to-right" } }]));
assert.ok(columnText.diagnostics.some(d => d.path.endsWith(".style.columnDirection") && d.status !== "supported"));
const verticalText = assessPpjPreviewInput(deck([{ ...text, style: { verticalText: "horizontal" } }]));
assert.ok(verticalText.diagnostics.some(d => d.path.endsWith(".style.verticalText") && d.status !== "supported"));
const unwrappedText = assessPpjPreviewInput(deck([{ ...text, style: { wrap: "none" } }]));
assert.ok(unwrappedText.diagnostics.some(d => d.path.endsWith(".style.wrap") && d.status !== "supported"));
const clippedText = assessPpjPreviewInput(deck([{ ...text, style: { horizontalOverflow: "clip" } }]));
assert.ok(clippedText.diagnostics.some(d => d.path.endsWith(".style.horizontalOverflow") && d.status !== "supported"));
const ellipsisText = assessPpjPreviewInput(deck([{ ...text, style: { verticalOverflow: "ellipsis" } }]));
assert.ok(ellipsisText.diagnostics.some(d => d.path.endsWith(".style.verticalOverflow") && d.status !== "supported"));
const middleText = assessPpjPreviewInput(deck([{ ...text, style: { verticalAlignment: "middle" } }]));
assert.ok(middleText.diagnostics.some(d => d.path.endsWith(".style.verticalAlignment") && d.status !== "supported"));
const zeroGapText = assessPpjPreviewInput(deck([{ ...text, style: { columnGap: 0 } }]));
assert.ok(zeroGapText.diagnostics.some(d => d.path.endsWith(".style.columnGap") && d.status !== "supported"));
for (const edge of ["left", "top", "right", "bottom"]) {
  const zeroMarginText = assessPpjPreviewInput(deck([{ ...text, style: { margins: { [edge]: 0 } } }]));
  assert.ok(zeroMarginText.diagnostics.some(d => d.path.includes(".style.margins") && d.status !== "supported"));
}
for (const style of [{ autoFit: "none" }, { autoFit: "shrink-text", normalAutoFit: { lineSpacingReduction: 0 } }]) {
  const autoFitText = assessPpjPreviewInput(deck([{ ...text, style }]));
  assert.ok(autoFitText.diagnostics.some(d => d.path.includes(".style.autoFit") && d.status !== "supported"));
  if (style.normalAutoFit) assert.ok(autoFitText.diagnostics.some(d => d.path.includes(".style.normalAutoFit") && d.status !== "supported"));
}
const falseAnchorCenterText = assessPpjPreviewInput(deck([{ ...text, style: { anchorCenter: false } }]));
assert.ok(falseAnchorCenterText.diagnostics.some(d => d.path.endsWith(".style.anchorCenter") && d.status !== "supported"));
const falseAntiAliasText = assessPpjPreviewInput(deck([{ ...text, style: { forceAntiAlias: false } }]));
assert.ok(falseAntiAliasText.diagnostics.some(d => d.path.endsWith(".style.forceAntiAlias") && d.status !== "supported"));
const falseParagraphSpacingText = assessPpjPreviewInput(deck([{ ...text, style: { spaceFirstLastParagraph: false } }]));
assert.ok(falseParagraphSpacingText.diagnostics.some(d => d.path.endsWith(".style.spaceFirstLastParagraph") && d.status !== "supported"));
const falseCompatibleSpacingText = assessPpjPreviewInput(deck([{ ...text, style: { compatibleLineSpacing: false } }]));
assert.ok(falseCompatibleSpacingText.diagnostics.some(d => d.path.endsWith(".style.compatibleLineSpacing") && d.status !== "supported"));
const reflection = { blur: 0, distance: 0, angle: 0, startOpacity: 0, endOpacity: 1,
  startPosition: 0, endPosition: 1, fadeAngle: 0, scaleX: 0, scaleY: -1,
  skewX: 0, skewY: 0, alignment: "ctr", rotateWithShape: false };
const reflectionDefaults = assessPpjPreviewInput(deck([{ ...text, text: { paragraphs: [
  { style: { defaultText: { reflection } }, runs: [{ text: "Reflection defaults" }] }] } }]));
for (const field of Object.keys(reflection))
  assert.ok(reflectionDefaults.diagnostics.some(d => d.path.endsWith(".style.defaultText.reflection." + field) && d.status !== "supported"));
const paragraphShadow = { color: "#112233", opacity: 0, blur: 0, distance: 0, angle: 0,
  scaleX: 0, scaleY: -1, skewX: 0, skewY: 0, alignment: "ctr", rotateWithShape: false };
const shadowDefaults = assessPpjPreviewInput(deck([{ ...text, text: { paragraphs: [
  { style: { defaultText: { shadow: paragraphShadow } }, runs: [{ text: "Shadow defaults" }] }] } }]));
for (const field of Object.keys(paragraphShadow))
  assert.ok(shadowDefaults.diagnostics.some(d => d.path.endsWith(".style.defaultText.shadow." + field) && d.status !== "supported"));
for (const radius of [0, 2]) {
  const softEdgeDefaults = assessPpjPreviewInput(deck([{ ...text, text: { paragraphs: [
    { style: { defaultText: { softEdge: { radius } } }, runs: [{ text: "Soft-edge defaults" }] }] } }]));
  assert.ok(softEdgeDefaults.diagnostics.some(d => d.path.endsWith(".style.defaultText.softEdge.radius") && d.status !== "supported"));
}
for (const alignment of ["left", "center", "right", "justify", "distributed", "justifyLow", "thaiDistributed"]) {
  const input = deck([{ ...text, text: { paragraphs: [
    { style: { alignment }, runs: [{ text: "Paragraph alignment" }] }] } }]);
  const paragraph = assessPpjPreviewInput(input);
  assert.ok(paragraph.diagnostics.some(d => d.path.endsWith(".style.alignment") && d.status !== "supported"));
  if (["justifyLow", "thaiDistributed"].includes(alignment)) {
    const receipt = previewSceneFixture(input, [{ id: "p1", elements: [nativeElement("text", "shape", {
      ...emuFrame(1, 2, 100, 60), geometry: "textbox", text: "Paragraph alignment",
      textBody: { paragraphs: [{ alignment, runs: [{ content: { case: "text", value: "Paragraph alignment" } }] }] },
    })] }], ["$.pages[0].elements[0]"]);
    const painted = paintPpjSceneSvg(receipt);
    assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-alignment" && d.status === "partial"));
    assert.match(painted.pages[0].svg, /Paragraph alignment/);
  }
}
for (const level of [0, 8]) {
  const input = deck([{ ...text, text: { paragraphs: [
    { style: { level }, runs: [{ text: "Paragraph level" }] }] } }]);
  const paragraph = assessPpjPreviewInput(input);
  assert.ok(paragraph.diagnostics.some(d => d.path.endsWith(".style.level") && d.status !== "supported"));
  const receipt = previewSceneFixture(input, [{ id: "p1", elements: [nativeElement("text", "shape", {
    ...emuFrame(1, 2, 100, 60), geometry: "textbox", text: "Paragraph level",
    textBody: { paragraphs: [{ level, runs: [{ content: { case: "text", value: "Paragraph level" } }] }] },
  })] }], ["$.pages[0].elements[0]"]);
  const painted = paintPpjSceneSvg(receipt);
  assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith(".level") && d.status === "partial"));
  assert.match(painted.pages[0].svg, /Paragraph level/);
}
for (const startAt of [1, 32767]) {
  const input = deck([{ ...text, text: { paragraphs: [
    { style: { bullet: { type: "number", scheme: "arabicPeriod", startAt } }, runs: [{ text: "Numbered paragraph" }] }] } }]);
  assert.ok(assessPpjPreviewInput(input).diagnostics.some(d => d.path.endsWith(".style.bullet.startAt") && d.status !== "supported"));
  const receipt = previewSceneFixture(input, [{ id: "p1", elements: [nativeElement("text", "shape", {
    ...emuFrame(1, 2, 100, 60), geometry: "textbox", text: "Numbered paragraph",
    textBody: { paragraphs: [{ bullet: { case: "autoNumber", value: { scheme: "arabicPeriod", startAt } },
      runs: [{ content: { case: "text", value: "Numbered paragraph" } }] }] },
  })] }], ["$.pages[0].elements[0]"]);
  const painted = paintPpjSceneSvg(receipt);
  assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith(".autoNumber") && d.status === "partial"));
  assert.match(painted.pages[0].svg, /Numbered paragraph/);
}
for (const [field, value, scheme] of [["scheme", "thaiNumParenBoth", "thaiNumParenBoth"], ["format", "upper-roman", "romanUcPeriod"]]) {
  const input = deck([{ ...text, text: { paragraphs: [
    { style: { bullet: { type: "number", [field]: value } }, runs: [{ text: "Numbering format" }] }] } }]);
  assert.ok(assessPpjPreviewInput(input).diagnostics.some(d => d.path.endsWith(".style.bullet." + field) && d.status !== "supported"));
  const receipt = previewSceneFixture(input, [{ id: "p1", elements: [nativeElement("text", "shape", {
    ...emuFrame(1, 2, 100, 60), geometry: "textbox", text: "Numbering format",
    textBody: { paragraphs: [{ bullet: { case: "autoNumber", value: { scheme } },
      runs: [{ content: { case: "text", value: "Numbering format" } }] }] },
  })] }], ["$.pages[0].elements[0]"]);
  const painted = paintPpjSceneSvg(receipt);
  assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.unmapped" && d.scenePath.endsWith(".autoNumber") && d.status === "partial"));
  assert.match(painted.pages[0].svg, /Numbering format/);
}
for (const character of ["★", "😀", "&"]) {
  const input = deck([{ ...text, text: { paragraphs: [{ style: { indent: 20, hanging: 10,
    bullet: { type: "character", character, fontFamily: "Georgia", size: 12, color: "#112233" } },
    runs: [{ text: "Character bullet" }] }] } }]);
  assert.ok(assessPpjPreviewInput(input).diagnostics.some(d => d.path.endsWith(".style.bullet.character") && d.status !== "supported"));
  const receipt = (font = true) => previewSceneFixture(input, [{ id: "p1", elements: [nativeElement("text", "shape", {
    ...emuFrame(1, 2, 100, 60), geometry: "textbox", text: "Character bullet",
    textBody: { paragraphs: [{ bullet: { case: "bulletCharacter", value: character },
      ...(font ? { bulletFont: { case: "bulletFontFamily", value: "Georgia" } } : {}),
      bulletColor: { case: "bulletColorRgb", value: "112233" }, bulletSize: { case: "bulletSizePoints", value: 12 },
      leftMargin: { case: "marginLeftEmu", value: 254000n }, indentation: { case: "indentEmu", value: -127000n },
      runs: [{ content: { case: "text", value: "Character bullet" } }] }] },
  })] }], ["$.pages[0].elements[0]"]);
  const painted = paintPpjSceneSvg(receipt());
  assert.match(painted.pages[0].svg, /data-officekit-bullet="character"/);
  assert.ok(painted.pages[0].svg.includes(`>${character === "&" ? "&amp;" : character}</text>`));
  assert.ok(painted.diagnostics.some(d => d.reason === "preview.scene.paint.text-layout" && d.status === "partial"));
  assert.ok(!painted.diagnostics.some(d => d.reason === "preview.scene.paint.bullet-layout"));
  assert.ok(paintPpjSceneSvg(receipt(false)).diagnostics.some(d => d.reason === "preview.scene.paint.bullet-layout" && d.status === "unavailable"));
}
for (const field of ["spaceBefore", "spaceBeforeMultiplier", "spaceAfter", "spaceAfterMultiplier", "lineSpacing", "lineSpacingMultiplier", "indent", "hanging"]) {
  const spacing = assessPpjPreviewInput(deck([{ ...text, text: { paragraphs: [
    { style: { [field]: field.startsWith("lineSpacing") ? 1 : 0 }, runs: [{ text: "Paragraph spacing" }] }] } }]));
  assert.ok(spacing.diagnostics.some(d => d.path.endsWith(".style." + field) && d.status !== "supported"));
}
for (const field of ["bold", "italic", "size", "fontFamily", "fontFamilyEastAsia", "fontFamilyComplexScript", "language", "kerning", "letterSpacing", "baseline", "capitalization", "strike", "underline", "highlight", "color", "gradient", "glow", "innerShadow"]) {
  const defaults = assessPpjPreviewInput(deck([{ ...text, text: { paragraphs: [{ style: { defaultText: { [field]: field === "innerShadow" ? { color: "#112233", blur: 0, distance: 0, angle: 0, opacity: 0 } : field === "glow" ? { color: "#112233", radius: 0, opacity: 0 } : field === "gradient" ? { kind: "linear", angle: 45, stops: [{ offset: 0, color: "#112233", opacity: 0 }, { offset: 1, color: "#FFFFFF", opacity: 1 }] } : ["highlight", "color"].includes(field) ? "#FFFF00" : field.startsWith("fontFamily") ? "Georgia" : ["size", "kerning", "letterSpacing", "baseline"].includes(field) ? 18.25 : field === "language" ? "fr-FR" : ["capitalization", "underline"].includes(field) ? "none" : false } }, runs: [{ text: "Default style" }] }] } }]));
  if (field === "gradient") {
    for (const leaf of ["angle", "kind", "stops[0].offset", "stops[0].color", "stops[0].opacity", "stops[1].opacity"])
      assert.ok(defaults.diagnostics.some(d => d.path.endsWith(".style.defaultText.gradient." + leaf) && d.status !== "supported"));
  } else if (field === "innerShadow") {
    for (const leaf of ["color", "blur", "distance", "angle", "opacity"])
      assert.ok(defaults.diagnostics.some(d => d.path.endsWith(".style.defaultText.innerShadow." + leaf) && d.status !== "supported"));
  } else if (field === "glow") {
    for (const leaf of ["color", "radius", "opacity"])
      assert.ok(defaults.diagnostics.some(d => d.path.endsWith(".style.defaultText.glow." + leaf) && d.status !== "supported"));
  } else assert.ok(defaults.diagnostics.some(d => d.path.endsWith(".style.defaultText." + field) && d.status !== "supported"));
}
const explicitNoWarp = assessPpjPreviewInput(deck([{ ...text, style: { textWarpPreset: "textNoShape", textWarpAdjustments: [] } }]));
assert.ok(explicitNoWarp.diagnostics.some(d => d.path.endsWith(".style.textWarpPreset") && d.status !== "supported"));
const zeroFlatText = assessPpjPreviewInput(deck([{ ...text, style: { flatTextZ: 0 } }]));
assert.ok(zeroFlatText.diagnostics.some(d => d.path.endsWith(".style.flatTextZ") && d.status !== "supported"));
const falseWordArtText = assessPpjPreviewInput(deck([{ ...text, style: { fromWordArt: false } }]));
assert.ok(falseWordArtText.diagnostics.some(d => d.path.endsWith(".style.fromWordArt") && d.status !== "supported"));
const singleColumnText = assessPpjPreviewInput(deck([{ ...text, style: { columns: 1 } }]));
assert.ok(singleColumnText.diagnostics.some(d => d.path.endsWith(".style.columns") && d.status !== "supported"));
const rotatedText = assessPpjPreviewInput(deck([{ ...text, style: { rotation: 0 } }]));
assert.ok(rotatedText.diagnostics.some(d => d.path.endsWith(".style.rotation") && d.status !== "supported"));

const rectangleShape = assessPpjPreviewInput(deck([{ type: "shape", id: "custom", frame, text: "Text",
  geometry: { kind: "custom", viewBox: frame, paths: [],
    textRectangle: { left: 10, top: 5, right: "r", bottom: "b" } } }]));
for (const edge of ["left", "top", "right", "bottom"])
  assert.ok(rectangleShape.diagnostics.some(d => d.path.endsWith(`.geometry.textRectangle.${edge}`) && d.status !== "supported"));

const guidedShape = assessPpjPreviewInput(deck([{ type: "shape", id: "guided", frame,
  geometry: { kind: "custom", viewBox: frame, paths: [], guides: [{ name: "inset", formula: "*/ w 1 10" }],
    textRectangle: { left: "inset", top: 5, right: "r", bottom: "b" } } }]));
assert.ok(guidedShape.diagnostics.some(d => d.path.endsWith(".geometry.guides[0].formula") && d.status !== "supported"));

const adjustedCustomShape = assessPpjPreviewInput(deck([{ type: "shape", id: "adjusted-custom", frame,
  geometry: { kind: "custom", viewBox: frame, paths: [], adjustments: [{ name: "padding", formula: "val 127000" }] } }]));
assert.ok(adjustedCustomShape.diagnostics.some(d => d.path.endsWith(".geometry.adjustments[0].formula") && d.status !== "supported"));

const extrusionShape = assessPpjPreviewInput(deck([{ type: "shape", id: "extrusion", frame,
  geometry: { kind: "custom", viewBox: frame, paths: [{ extrusionOk: false, commands: [{ op: "moveTo", x: 0, y: 0 }] }] } }]));
assert.ok(extrusionShape.diagnostics.some(d => d.path.endsWith(".geometry.paths[0].extrusionOk") && d.status !== "supported"));

const siteShape = assessPpjPreviewInput(deck([{ type: "shape", id: "site", frame,
  geometry: { kind: "custom", viewBox: frame, paths: [], connectionSites: [{ angle: 90, x: "hc", y: 10 }] } }]));
assert.ok(siteShape.diagnostics.some(d => d.path.includes(".geometry.connectionSites[0]") && d.status !== "supported"));

const handleShape = assessPpjPreviewInput(deck([{ type: "shape", id: "handle", frame,
  geometry: { kind: "custom", viewBox: frame, paths: [], adjustmentHandles: [{ kind: "xy", xAdjustment: "ax", position: { x: "ax", y: 10 } }] } }]));
assert.ok(handleShape.diagnostics.some(d => d.path.includes(".geometry.adjustmentHandles[0]") && d.status !== "supported"));

const referencePathShape = assessPpjPreviewInput(deck([{ type: "shape", id: "reference-path", frame,
  geometry: { kind: "custom", viewBox: frame, paths: [{ commands: [{ op: "moveTo", x: "hc", y: 10 }] }] } }]));
assert.ok(referencePathShape.diagnostics.some(d => d.path.endsWith(".geometry.paths[0].commands[0].x") && d.status !== "supported"));

const viewportShape = assessPpjPreviewInput(deck([{ type: "shape", id: "viewport", frame,
  geometry: { kind: "custom", viewBox: frame, paths: [{ viewport: { width: 0, height: 50 }, commands: [{ op: "moveTo", x: 0, y: 0 }] }] } }]));
assert.ok(viewportShape.diagnostics.some(d => d.path.includes(".geometry.paths[0].viewport") && d.status !== "supported"));

const repeated = { pages: [{ id: "p1", elements: [text] }, { id: "p2", elements: [text] }] };
const diagnostics = assessPpjPreviewInput(repeated).diagnostics.filter((d) => d.path.endsWith(".text") && d.reason === "preview.text.unassessed");
assert.deepEqual(diagnostics.map((d) => d.pageId).sort(), ["p1", "p2"]);

const metadata = { ...deck([]), meta: { id: "deck", title: "title", version: 1, language: "en" } };
assert.ok(!assessPpjPreviewInput(metadata).diagnostics.some((d) => d.path.startsWith("$.meta")));
metadata.meta.newVisualField = { opacity: 0, hidden: false, missing: null };
const unknown = assessPpjPreviewInput(metadata);
for (const [key, expected] of [["opacity", "0"], ["hidden", "false"], ["missing", "null"]]) {
  const d = unknown.diagnostics.find((item) => item.path === `$.meta.newVisualField.${key}`);
  assert.equal(d.reason, "preview.field.unassessed"); assert.equal(d.valueSummary, expected);
}
assert.equal(unknown.reliability.status, "requires-review");
for (const key of ["theme", "styles", "masters", "layouts", "grammar"]) {
  const state = { ...deck([text]), design: { [key]: key === "masters" || key === "layouts" ? [{ id: "layout" }] : { unknown: "test" } } };
  const assessed = assessPpjPreviewInput(state);
  const leaf = assessed.children[0].children[0];
  assert.ok(leaf.diagnostics.some((d) => d.path.startsWith(`$.design.${key}`)), key);
  assert.notEqual(leaf.status, "supported");
}
const component = deck([{ id: "component", type: "component", frame, component: "definition", arguments: { value: 0 }, slots: { body: [text] }, repeat: { items: [{ key: "one", arguments: { value: 0 } }] } }]);
component.components = [{ id: "definition", frame, elements: [text] }];
const c = assessPpjPreviewInput(component);
assert.ok(c.diagnostics.some((d) => d.path === "$.components[0].elements[0].text"));
assert.ok(c.diagnostics.some((d) => d.path === "$.pages[0].elements[0].slots.body[0].text" && d.id === "text"));
assert.ok(c.diagnostics.some((d) => d.path.endsWith(".arguments.value") && d.valueSummary === "0"));
assert.ok(c.children.find((item) => item.pageId === "p1").children[0].diagnostics.some((d) => d.path === "$.components[0].elements[0].text"));
const transformed = structuredClone(sample); transformed.pages[0].elements[0].frame.rotation = 30;
const nested = assessPpjPreviewInput(transformed).children[0].children[0].children[0];
assert.ok(nested.diagnostics.some((d) => d.path === "$.pages[0].elements[0].frame.rotation" && d.id === "group"));

const source = { hiddenPayload: "secret native payload" };
Object.defineProperty(source, "mustNotRead", { enumerable: true, get() { throw new Error("Source payload was expanded"); } });
const native = deck([{ id: "native", type: "opaque", frame, nativeKind: "chart", summary: "native", nativeRef: source }]);
native.source = source;
const retained = assessPpjPreviewInput(native);
assert.equal(retained.status, "opaque");
assert.ok(retained.diagnostics.some((d) => d.path.endsWith(".nativeRef") && d.valueSummary === "[source-owned value omitted]"));
assert.ok(!JSON.stringify(retained).includes("secret native payload"));
assert.ok(!retained.diagnostics.some((d) => d.path.includes("hiddenPayload")));
const unexpectedBinary = deck([{ ...text, extra: new Uint8Array([65, 66]) }]);
assert.ok(assessPpjPreviewInput(unexpectedBinary).diagnostics.some((d) => d.path.endsWith(".extra") && d.valueSummary.includes("omitted")));
const cyclic = deck([text]); cyclic.extra = cyclic;
assert.equal(assessPpjPreviewInput(cyclic).reliability.status, "failed");

// The real canonical program exercises source schema, metadata and nested
// component definitions. It remains review-required; no drawing is claimed.
const canonical = JSON.parse(await readFile("test/fixtures/presentation/evidence-ledger-canonical.ppj", "utf8"));
const bytes = JSON.stringify(canonical);
const actual = assessPpjPreviewInput(canonical);
assert.equal(JSON.stringify(canonical), bytes);
assert.notEqual(actual.status, "supported");
assert.ok(actual.diagnostics.some((d) => d.path.includes(".text")));
assert.ok(actual.diagnostics.every((d) => d.path.startsWith("$") && d.reason && d.action && d.valueSummary.length <= 256));
console.log("ppj preview input assessment ok: schema fields, nesting, metadata, inheritance and source preservation");
