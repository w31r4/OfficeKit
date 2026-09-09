// Pure preview evidence helpers. No filesystem, native codec or raster imports.
// Support describes the assessed visual state, not file production or host QA.
export const PREVIEW_STATUSES = Object.freeze(["supported", "partial", "opaque", "unavailable"]);
export const PREVIEW_SEVERITIES = Object.freeze(["info", "warning", "error"]);
const ranks = new Map(PREVIEW_STATUSES.map((status, index) => [status, index]));
const SUMMARY_LIMIT = 256;

function nonempty(value, name) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`Preview ${name} must be a nonempty string`);
  return value;
}

function validStatus(status) {
  if (!ranks.has(status)) throw new TypeError(`Invalid preview status: ${String(status)}`);
  return status;
}

function validateIdentity({ path, pageId, id, scenePath }) {
  nonempty(path, "path");
  if (path !== "$" && !path.startsWith("$.") && !path.startsWith("$[")) throw new TypeError("Preview path must be rooted at $");
  if (scenePath !== undefined) {
    nonempty(scenePath, "scene path");
    if (scenePath !== "$" && !scenePath.startsWith("$.") && !scenePath.startsWith("$["))
      throw new TypeError("Preview scene path must be rooted at $");
  }
  if (pageId !== undefined) nonempty(pageId, "page ID");
  if (id !== undefined) nonempty(id, "element ID");
  if (id !== undefined && pageId === undefined) throw new TypeError("Preview element identity requires a page ID");
}

function validateDiagnostic(diagnostic) {
  const { status, severity } = diagnostic;
  validStatus(status);
  validateIdentity(diagnostic);
  nonempty(diagnostic.reason, "reason");
  nonempty(diagnostic.action, "action");
  if (!PREVIEW_SEVERITIES.includes(severity)) throw new TypeError(`Invalid preview severity: ${String(severity)}`);
  if (status === "supported" && severity !== "info") throw new TypeError("Supported preview evidence must be informational");
  if (status !== "supported" && severity === "info") throw new TypeError("A preview limitation cannot be informational");
  if (status === "unavailable" && severity !== "error") throw new TypeError("Unavailable preview evidence must be an error");
}

/** Stable address. Bracket notation keeps dots, brackets and quotes in IDs/keys unambiguous. */
export function previewPath(parent, key) {
  nonempty(parent, "parent path");
  if (Number.isSafeInteger(key) && key >= 0) return `${parent}[${key}]`;
  if (typeof key !== "string") throw new TypeError("Preview path key must be a string or nonnegative safe integer");
  return /^[A-Za-z_$][\w$]*$/u.test(key) ? `${parent}.${key}` : `${parent}[${JSON.stringify(key)}]`;
}

/** XML escaping is deliberately separate from JSON evidence serialization. */
export function escapePreviewText(value) {
  return String(value).replace(/[&<>"']/gu, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[c]);
}

// Summaries never enumerate objects or visit array items: large/cyclic payloads,
// getters and source package bytes cannot expand into a diagnostic by accident.
export function previewValueSummary(value, { sensitive = false } = {}) {
  if (sensitive) return "[source-owned value omitted]";
  if (value === null) return "null";
  if (value === undefined) return "[not supplied]";
  if (ArrayBuffer.isView(value) || value instanceof ArrayBuffer) return "[binary value omitted]";
  if (Array.isArray(value)) return `[array: ${value.length} items]`;
  if (typeof value === "object") return "[object; inspect the input path]";
  if (typeof value === "function" || typeof value === "symbol") return `[${typeof value} value omitted]`;
  const raw = String(value);
  // Encode XML-sensitive/control characters while bounding the encoded result.
  // Do not cut an entity in half or allocate an encoded copy of a huge string.
  let summary = "";
  for (const character of raw) {
    const encoded = /[\u0000-\u001f\u007f-\u009f]/u.test(character)
      ? `\\u${character.codePointAt(0).toString(16).padStart(4, "0")}` : escapePreviewText(character);
    if (summary.length + encoded.length > SUMMARY_LIMIT - 1) return `${summary}…`;
    summary += encoded;
  }
  return summary;
}

/** Create a validated, serializable diagnostic without retaining the input value. */
export function previewDiagnostic({
  pageId, id, path, scenePath, status = "partial", reason, severity,
  value, sensitive = false, action,
}) {
  validStatus(status);
  severity ??= status === "supported" ? "info" : status === "unavailable" ? "error" : "warning";
  validateDiagnostic({ pageId, id, path, scenePath, status, reason, severity, action });
  return Object.freeze({
    ...(pageId === undefined ? {} : { pageId }), ...(id === undefined ? {} : { id }),
    ...(scenePath === undefined ? {} : { scenePath }),
    path, status, reason, severity, valueSummary: previewValueSummary(value, { sensitive }), action,
  });
}

/** Empty/unassessed state is partial, not a silent successful assessment. */
export function aggregatePreviewStatus(statuses, { assessed = false } = {}) {
  let result = assessed ? "supported" : "partial";
  for (const status of statuses) {
    validStatus(status);
    if (ranks.get(status) > ranks.get(result)) result = status;
  }
  return result;
}

function diagnosticKey(diagnostic) {
  // One semantic owner can lower to many native nodes and fields. Their
  // diagnostics must not collapse merely because they share an input path.
  return JSON.stringify([diagnostic.pageId ?? null, diagnostic.id ?? null, diagnostic.path, diagnostic.scenePath ?? null, diagnostic.reason]);
}

/** Deduplicate identities without allowing a weak duplicate to hide an error. */
export function mergePreviewDiagnostics(...groups) {
  const byKey = new Map();
  for (const diagnostic of groups.flat()) {
    validateDiagnostic(diagnostic);
    const key = diagnosticKey(diagnostic), previous = byKey.get(key);
    if (!previous) byKey.set(key, Object.freeze({ ...diagnostic }));
    else {
      const severity = PREVIEW_SEVERITIES[Math.max(PREVIEW_SEVERITIES.indexOf(previous.severity), PREVIEW_SEVERITIES.indexOf(diagnostic.severity))];
      const status = aggregatePreviewStatus([previous.status, diagnostic.status], { assessed: true });
      // Stable tie-breaking makes union independent of traversal order.
      const weight = (item) => PREVIEW_SEVERITIES.indexOf(item.severity) * PREVIEW_STATUSES.length + ranks.get(item.status);
      const selected = weight(previous) > weight(diagnostic) ? previous : weight(diagnostic) > weight(previous) ? diagnostic
        : JSON.stringify(previous) <= JSON.stringify(diagnostic) ? previous : diagnostic;
      byKey.set(key, Object.freeze({ ...selected, status, severity }));
    }
  }
  return Object.freeze([...byKey.entries()].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0).map(([, diagnostic]) => diagnostic));
}

export function previewReliability(diagnostics, status) {
  validStatus(status);
  for (const diagnostic of diagnostics) validateDiagnostic(diagnostic);
  status = aggregatePreviewStatus([status, ...diagnostics.map((diagnostic) => diagnostic.status)], { assessed: true });
  const violations = diagnostics.filter((diagnostic) => diagnostic.severity === "error" || diagnostic.status === "unavailable");
  return Object.freeze({
    status: violations.length || status === "unavailable" ? "failed"
      : status !== "supported" ? "requires-review" : "passed",
    violations: Object.freeze(violations),
  });
}

/** Same constructor for an element, page or document; children keep their own evidence. */
export function previewAssessment({ path = "$", scenePath, pageId, id, assessed = false, diagnostics = [], children = [] } = {}) {
  validateIdentity({ path, scenePath, pageId, id });
  const own = assessed ? diagnostics : [...diagnostics, previewDiagnostic({
    pageId, id, path, scenePath, reason: "preview.state.unassessed", value: undefined,
    action: "Assess this input state before using the preview as visual evidence.",
  })];
  const all = mergePreviewDiagnostics(own, ...children.map((child) => child.diagnostics));
  const status = aggregatePreviewStatus([...all.map((diagnostic) => diagnostic.status), ...children.map((child) => child.status)], { assessed });
  return Object.freeze({
    ...(pageId === undefined ? {} : { pageId }), ...(id === undefined ? {} : { id }),
    ...(scenePath === undefined ? {} : { scenePath }),
    path, status, diagnostics: all, reliability: previewReliability(all, status),
    children: Object.freeze([...children]),
  });
}
