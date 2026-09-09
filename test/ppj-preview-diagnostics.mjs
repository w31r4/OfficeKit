import assert from "node:assert/strict";
import {
  PREVIEW_STATUSES, PREVIEW_SEVERITIES, previewPath, escapePreviewText,
  previewValueSummary, previewDiagnostic, aggregatePreviewStatus,
  mergePreviewDiagnostics, previewReliability, previewAssessment,
} from "../src/ppj/preview-diagnostics.mjs";

const make = (options = {}) => previewDiagnostic({
  pageId: "page-a", id: "node", path: "$.pages[0].elements[0].text", reason: "preview.text.unassessed",
  value: "hello", action: "Inspect this text field against the input.", ...options,
});
for (const status of PREVIEW_STATUSES) {
  const diagnostic = make({ status });
  assert.equal(diagnostic.status, status);
  assert.equal(diagnostic.severity, status === "supported" ? "info" : status === "unavailable" ? "error" : "warning");
  assert.equal(diagnostic.path, "$.pages[0].elements[0].text");
  assert.ok(Object.isFrozen(diagnostic));
  assert.ok(!Object.hasOwn(diagnostic, "value"));
}
for (const options of [
  { status: "full" }, { status: "__proto__" }, { severity: "fatal" },
  { status: "supported", severity: "error" }, { status: "partial", severity: "info" },
  { status: "unavailable", severity: "warning" }, { path: "pages[0]" }, { path: "" },
  { reason: "" }, { action: "" }, { pageId: "" }, { id: "" }, { pageId: undefined },
]) assert.throws(() => make(options), TypeError);
assert.deepEqual(PREVIEW_SEVERITIES, ["info", "warning", "error"]);
assert.equal(previewPath(previewPath(previewPath("$", "pages"), 0), "strange.key\""), '$.pages[0]["strange.key\\\""]');
assert.equal(previewPath("$", "__proto__"), "$.__proto__");
assert.throws(() => previewPath("$", -1), TypeError);
assert.throws(() => previewPath("$", NaN), TypeError);

assert.equal(previewValueSummary(null), "null");
assert.equal(previewValueSummary(false), "false");
assert.equal(previewValueSummary(0), "0");
assert.equal(previewValueSummary(undefined), "[not supplied]");
assert.equal(previewValueSummary(""), "");
assert.equal(previewValueSummary(new Uint8Array([1, 2])), "[binary value omitted]");
assert.equal(previewValueSummary(new ArrayBuffer(100)), "[binary value omitted]");
assert.equal(previewValueSummary("secret", { sensitive: true }), "[source-owned value omitted]");
const cyclic = {}; cyclic.self = cyclic;
Object.defineProperty(cyclic, "getter", { get() { throw new Error("Must not inspect object contents"); }, enumerable: true });
assert.equal(previewValueSummary(cyclic), "[object; inspect the input path]");
assert.equal(previewValueSummary([cyclic]), "[array: 1 items]");
assert.equal(previewValueSummary('<script a="\u0000">&\n'), "&lt;script a=&quot;\\u0000&quot;&gt;&amp;\\u000a");
for (const raw of ["x".repeat(1000000), "<&\u0000".repeat(10000), "😀".repeat(10000)]) {
  const summary = previewValueSummary(raw);
  assert.ok(summary.length <= 256);
  assert.ok(summary.endsWith("…"));
  assert.doesNotMatch(summary, /[<>\u0000-\u001f]/u);
  assert.doesNotMatch(summary, /&[^;]*…$/u, "XML entity must not be cut in half");
}
assert.equal(escapePreviewText('<&"\''), "&lt;&amp;&quot;&apos;");

for (const a of PREVIEW_STATUSES) for (const b of PREVIEW_STATUSES) {
  const expected = PREVIEW_STATUSES[Math.max(PREVIEW_STATUSES.indexOf(a), PREVIEW_STATUSES.indexOf(b))];
  assert.equal(aggregatePreviewStatus([a, b], { assessed: true }), expected);
  assert.equal(aggregatePreviewStatus([b, a], { assessed: true }), expected);
}
assert.equal(aggregatePreviewStatus([]), "partial");
assert.equal(aggregatePreviewStatus([], { assessed: true }), "supported");
assert.equal(aggregatePreviewStatus(["supported"]), "partial");
assert.throws(() => aggregatePreviewStatus(["full"]), TypeError);

const info = make({ status: "supported" });
const warning = make();
const failure = make({ severity: "error" });
const missing = make({ status: "unavailable", reason: "preview.asset.missing" });
assert.deepEqual(mergePreviewDiagnostics([warning, failure, info]), mergePreviewDiagnostics([failure, info, warning]));
assert.equal(mergePreviewDiagnostics([warning, failure]).length, 1);
assert.equal(mergePreviewDiagnostics([warning, failure])[0].severity, "error");
assert.equal(mergePreviewDiagnostics([warning, failure])[0].status, "partial");
assert.equal(mergePreviewDiagnostics([warning, make({ pageId: "page-b" })]).length, 2);
assert.equal(mergePreviewDiagnostics([warning, make({ path: "$.pages[0].elements[1].text" })]).length, 2);
assert.equal(previewReliability([failure], "partial").status, "failed");
assert.equal(previewReliability([missing], "supported").status, "failed");
assert.equal(previewReliability([warning], "partial").status, "requires-review");
assert.equal(previewReliability([warning], "supported").status, "requires-review", "a caller cannot upgrade diagnostic evidence");
assert.equal(previewReliability([info], "supported").status, "passed");
assert.throws(() => mergePreviewDiagnostics([{ ...warning, severity: "info" }]), TypeError);
assert.throws(() => previewAssessment({ assessed: true, path: "invalid" }), TypeError);

const unknown = previewAssessment();
assert.equal(unknown.status, "partial");
assert.equal(unknown.reliability.status, "requires-review");
assert.equal(unknown.diagnostics[0].reason, "preview.state.unassessed");
const supported = previewAssessment({ assessed: true, diagnostics: [info] });
assert.equal(supported.status, "supported", "informational evidence must not downgrade support");
const element = previewAssessment({ path: "$.pages[0].elements[0]", pageId: "page-a", id: "node", assessed: true, diagnostics: [failure] });
const page = previewAssessment({ path: "$.pages[0]", pageId: "page-a", assessed: true, children: [supported, element] });
const document = previewAssessment({ assessed: true, children: [page] });
assert.equal(document.reliability.status, "failed");
assert.deepEqual(document.diagnostics, page.diagnostics);
assert.deepEqual(document.reliability.violations, element.reliability.violations);
assert.equal(document.children[0].children[1].id, "node");
assert.equal(JSON.parse(JSON.stringify(document)).reliability.status, "failed");
assert.ok(Object.isFrozen(document.children));
assert.equal(previewAssessment({ assessed: true, children: [page, previewAssessment({ assessed: true, diagnostics: [missing] })] }).status, "unavailable");
console.log("ppj preview diagnostic primitives ok: paths, bounded values, aggregation and reliability");
await import("./ppj-preview-input-assessment.mjs");
await import("./ppj-preview-field-limits.mjs");
await import("./ppj-preview-factual-errors.mjs");
await import("./ppj-preview-render-assessment.mjs");
await import("./ppj-preview-scene-wire.mjs");
await import("./ppj-preview-scene-transport.mjs");
