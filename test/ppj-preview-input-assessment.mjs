import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { assessPpjPreviewInput } from "../src/ppj/preview-input-assessment.mjs";

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
