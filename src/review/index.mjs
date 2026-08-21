import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { DocumentFile, DocumentModel } from "../document/index.mjs";
import { PdfArtifact, PdfFile } from "../pdf/index.mjs";
import { Presentation, PresentationFile } from "../presentation/index.mjs";
import { createArtifactVisualQaApi } from "../qa/artifact-visual.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { toUint8Array } from "../shared/binary.mjs";
import { SpreadsheetFile, Workbook } from "../spreadsheet/index.mjs";

const ANYDOC_VERSION = "0.1.3";
const DEFAULT_MAX_BYTES = 100 * 1024 * 1024;
const DEFAULT_MAX_CONTENT_CHARS = 40_000;
const DEFAULT_MAX_INSPECT_CHARS = 20_000;
const DEFAULT_MAX_SUMMARY_CHARS = 50_000;

const FORMAT_DETAILS = {
  docx: {
    artifactKind: "document",
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  },
  xlsx: {
    artifactKind: "workbook",
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  },
  pptx: {
    artifactKind: "presentation",
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  },
  pdf: { artifactKind: "pdf", type: "application/pdf" },
};

const FORMAT_ALIASES = new Map([
  ["document", "docx"],
  ["word", "docx"],
  ["docx", "docx"],
  ["workbook", "xlsx"],
  ["spreadsheet", "xlsx"],
  ["excel", "xlsx"],
  ["xlsx", "xlsx"],
  ["presentation", "pptx"],
  ["powerpoint", "pptx"],
  ["pptx", "pptx"],
  ["pdf", "pdf"],
]);

const MIME_FORMATS = new Map(Object.entries(FORMAT_DETAILS).map(([format, details]) => [details.type, format]));
const VISUAL_REVIEW_STATUSES = new Set(["complete", "unavailable", "requires-human"]);

function positiveInteger(value, fallback, label) {
  if (value == null) return fallback;
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number <= 0) throw new TypeError(`${label} must be a positive safe integer.`);
  return number;
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function modelFormat(value) {
  if (value instanceof DocumentModel) return "docx";
  if (value instanceof Workbook) return "xlsx";
  if (value instanceof Presentation) return "pptx";
  if (value instanceof PdfArtifact) return "pdf";
  return undefined;
}

function normalizeFormat(value) {
  if (value == null || value === "") return undefined;
  const normalized = String(value).trim().toLowerCase().replace(/^\./u, "");
  return FORMAT_ALIASES.get(normalized);
}

function formatFromPath(value) {
  return normalizeFormat(path.extname(value));
}

function formatFromInputMetadata(input) {
  return normalizeFormat(input?.metadata?.artifactKind)
    || normalizeFormat(input?.metadata?.format)
    || MIME_FORMATS.get(String(input?.type || "").toLowerCase());
}

function formatFromPdfSignature(bytes) {
  return bytes?.length >= 5
    && bytes[0] === 0x25
    && bytes[1] === 0x50
    && bytes[2] === 0x44
    && bytes[3] === 0x46
    && bytes[4] === 0x2d
    ? "pdf"
    : undefined;
}

function selectFormat(input, bytes, options = {}) {
  const candidates = [
    modelFormat(input),
    normalizeFormat(options.format),
    normalizeFormat(options.kind),
    typeof input === "string" ? formatFromPath(input) : undefined,
    formatFromInputMetadata(input),
    formatFromPdfSignature(bytes),
  ].filter(Boolean);
  const unique = [...new Set(candidates)];
  if (unique.length > 1) throw new TypeError(`Review input format declarations disagree: ${unique.join(", ")}.`);
  if (!unique.length) {
    throw new TypeError("Review input format is unknown. Use a .docx/.xlsx/.pptx/.pdf path, a typed FileBlob, a modeled artifact, or options.format.");
  }
  return unique[0];
}

async function readBoundedBytes(input, maxBytes, label) {
  if (typeof input === "string") {
    const absolutePath = path.resolve(input);
    const metadata = await fs.stat(absolutePath);
    if (!metadata.isFile()) throw new TypeError(`${label} must be a regular file.`);
    if (metadata.size > maxBytes) throw new RangeError(`${label} has ${metadata.size} bytes and exceeds maxBytes (${maxBytes}).`);
    const bytes = new Uint8Array(await fs.readFile(absolutePath));
    return { bytes, path: absolutePath };
  }
  let bytes;
  if (input instanceof FileBlob) bytes = input.bytes;
  else if (input instanceof Uint8Array || input instanceof ArrayBuffer || ArrayBuffer.isView(input)) bytes = toUint8Array(input);
  else if (typeof input?.arrayBuffer === "function") bytes = new Uint8Array(await input.arrayBuffer());
  else throw new TypeError(`${label} must be a local path, FileBlob, Uint8Array, ArrayBuffer, or Blob-like object.`);
  if (bytes.byteLength > maxBytes) throw new RangeError(`${label} has ${bytes.byteLength} bytes and exceeds maxBytes (${maxBytes}).`);
  return { bytes: new Uint8Array(bytes), path: undefined };
}

async function exportModel(model, format, options = {}) {
  if (format === "docx") return DocumentFile.exportDocx(model, options);
  if (format === "xlsx") return SpreadsheetFile.exportXlsx(model, options);
  if (format === "pptx") return PresentationFile.exportPptx(model, options);
  if (format === "pdf") return PdfFile.exportPdf(model, options);
  throw new TypeError(`Unsupported review format ${format}.`);
}

async function importModel(bytes, format, options = {}) {
  const blob = new FileBlob(bytes, { type: FORMAT_DETAILS[format].type });
  if (format === "docx") return DocumentFile.importDocx(blob, options);
  if (format === "xlsx") return SpreadsheetFile.importXlsx(blob, options);
  if (format === "pptx") return PresentationFile.importPptx(blob, options);
  if (format === "pdf") return PdfFile.importPdf(blob, options);
  throw new TypeError(`Unsupported review format ${format}.`);
}

async function inspectPackage(bytes, format, options = {}) {
  const blob = new FileBlob(bytes, { type: FORMAT_DETAILS[format].type });
  if (format === "docx") return DocumentFile.inspectDocx(blob, options);
  if (format === "xlsx") return SpreadsheetFile.inspectXlsx(blob, options);
  if (format === "pptx") return PresentationFile.inspectPptx(blob, options);
  if (format === "pdf") return PdfFile.inspectPdf(blob, options);
  throw new TypeError(`Unsupported review format ${format}.`);
}

function inferArtifactKind(artifact) {
  return artifact?.__officeKitReviewKind || FORMAT_DETAILS[modelFormat(artifact)]?.artifactKind || "unknown";
}

const visualQaApi = createArtifactVisualQaApi({ inferArtifactKind });

function boundedMessage(error, maxChars = 1_000) {
  return String(error?.message || error || "Unknown error").replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/gu, " ").slice(0, maxChars);
}

function parseNdjson(text) {
  return String(text || "").split(/\r?\n/u).filter(Boolean).flatMap((line) => {
    try { return [JSON.parse(line)]; }
    catch { return []; }
  });
}

function recordCounts(ndjsonText) {
  const counts = {};
  for (const record of parseNdjson(ndjsonText)) {
    if (record.kind === "notice") continue;
    const kind = String(record.kind || "unknown");
    counts[kind] = (counts[kind] || 0) + 1;
  }
  return counts;
}

function issueSeverity(issue) {
  return String(issue?.severity || "error").toLowerCase();
}

function hasHardIssue(issues = []) {
  return issues.some((issue) => issueSeverity(issue) === "error");
}

function reviewIssueFingerprint(issue) {
  return JSON.stringify({
    kind: issue?.kind,
    type: issue?.type,
    slide: issue?.slide,
    id: issue?.id,
    ids: issue?.ids,
    name: issue?.name,
    names: issue?.names,
    message: issue?.message,
  });
}

function applyBaselineReview(report, baselineReport) {
  // A source-bound baseline may explain unchanged modeled/layout findings,
  // but it never masks a structural package failure. The package itself must
  // still reopen and pass the independent OPC inspection.
  const sections = ["semantic", "layout"];
  let matchedIssues = 0;
  let newIssues = 0;
  for (const sectionName of sections) {
    const section = report[sectionName];
    const baseline = baselineReport?.[sectionName];
    if (!section || !baseline || !Array.isArray(section.issues) || !Array.isArray(baseline.issues)) continue;
    const baselineIssues = new Set(baseline.issues.map(reviewIssueFingerprint));
    section.issues = section.issues.map((issue) => {
      if (issueSeverity(issue) === "error" && baselineIssues.has(reviewIssueFingerprint(issue))) {
        matchedIssues += 1;
        return { ...issue, severity: "warning", preexisting: true };
      }
      if (issueSeverity(issue) === "error") newIssues += 1;
      return issue;
    });
    if (section.status === "skipped") continue;
    section.ok = !hasHardIssue(section.issues);
    section.status = section.ok
      ? section.issues.length ? "passed-with-warnings" : "passed"
      : "failed";
  }
  return { matchedIssues, newIssues };
}

function truncateMarkdown(value, maxChars) {
  const text = String(value || "").replace(/\r\n?/gu, "\n").trim();
  if (text.length <= maxChars) return { markdown: text, chars: text.length, originalChars: text.length, truncated: false };
  const notice = "\n\n[OfficeKit: content view truncated; narrow the task or raise maxContentChars.]";
  const limit = Math.max(0, maxChars - notice.length);
  let cut = text.slice(0, limit);
  const newline = cut.lastIndexOf("\n");
  if (newline >= Math.floor(limit * 0.7)) cut = cut.slice(0, newline);
  const markdown = `${cut.trimEnd()}${notice}`;
  return { markdown, chars: markdown.length, originalChars: text.length, truncated: true };
}

function removeEmbeddedDataUris(markdown) {
  return String(markdown || "").replace(/data:[a-z0-9.+-]+\/[a-z0-9.+-]+;base64,[a-z0-9+/=]+/giu, "officekit:embedded-image-bytes-omitted");
}

async function anyDocContentView(bytes, format, options = {}) {
  let anydoc;
  try {
    const module = await import("@firecrawl/anydoc");
    anydoc = module.toMarkdownBytes ? module : module.default;
    if (typeof anydoc?.toMarkdownBytes !== "function") throw new Error("toMarkdownBytes is unavailable");
  } catch (error) {
    return {
      status: "unavailable",
      requested: true,
      provider: "anydoc",
      providerVersion: ANYDOC_VERSION,
      format,
      reason: "provider-unavailable",
      message: boundedMessage(error),
    };
  }
  try {
    const markdown = removeEmbeddedDataUris(await anydoc.toMarkdownBytes(bytes, format));
    const bounded = truncateMarkdown(markdown, options.maxChars);
    return {
      status: bounded.markdown ? "ready" : "empty",
      requested: true,
      provider: "anydoc",
      providerVersion: ANYDOC_VERSION,
      format,
      sourceSha256: sha256(bytes),
      markdownSha256: sha256(new TextEncoder().encode(bounded.markdown)),
      ...bounded,
    };
  } catch (error) {
    return {
      status: "unavailable",
      requested: true,
      provider: "anydoc",
      providerVersion: ANYDOC_VERSION,
      format,
      sourceSha256: sha256(bytes),
      reason: /scan|image.only|ocr|unsupported/iu.test(boundedMessage(error)) ? "unsupported-content" : "conversion-failed",
      message: boundedMessage(error),
    };
  }
}

async function canonicalPath(value) {
  const absolute = path.resolve(value);
  try { return await fs.realpath(absolute); }
  catch (error) {
    if (error?.code !== "ENOENT") throw error;
    const parent = await fs.realpath(path.dirname(absolute));
    return path.join(parent, path.basename(absolute));
  }
}

function reviewIssue(type, message, severity = "error", details = {}) {
  return { kind: "reviewIssue", type, severity, message, ...details };
}

function summarizeIssues(issues = [], limit = 8) {
  if (!issues.length) return ["- No machine-detected issues."];
  const lines = issues.slice(0, limit).map((issue) => `- ${String(issue.severity || "error").toUpperCase()} ${issue.type || issue.kind || "issue"}: ${issue.message || "No message"}`);
  if (issues.length > limit) lines.push(`- ${issues.length - limit} more issue(s) omitted from this compact view.`);
  return lines;
}

function createReviewMarkdown(report, maxChars) {
  const semanticIssues = report.semantic.issues || [];
  const structuralIssues = report.structural.issues || [];
  const layoutIssues = report.layout.issues || [];
  const deliveryIssues = report.delivery.issues || [];
  const counts = Object.entries(report.semantic.recordCounts || {}).map(([kind, count]) => `${kind}=${count}`).join(", ") || "unavailable";
  const prefix = [
    "# OfficeKit post-edit review",
    "",
    `- Verdict: ${report.verdict}`,
    `- Artifact: ${report.artifactKind} (${report.delivery.type})`,
    `- SHA-256: ${report.delivery.sha256}`,
    `- Bytes: ${report.delivery.bytes}`,
    "",
    "## 6. Semantic review",
    "",
    `Status: ${report.semantic.status}. Modeled records: ${counts}.`,
    ...summarizeIssues(semanticIssues),
    "",
    "## 7. Structural review",
    "",
    `Status: ${report.structural.status}.`,
    ...summarizeIssues(structuralIssues),
    "",
    "## 8. Layout/render review",
    "",
    `Status: ${report.layout.status}${report.layout.scope ? `; scope: ${report.layout.scope}` : ""}.`,
    ...summarizeIssues(layoutIssues),
    "",
    "## 9. Text reading view (optional)",
    "",
    `Status: ${report.contentView.status}; provider: anydoc@${ANYDOC_VERSION}.`,
  ].join("\n");
  const suffix = [
    "## 10. Visual/human review",
    "",
    `Status: ${report.visualReview}. Text reading/OCR output does not count as pixel or aesthetic review.`,
    "",
    "## 11. Delivery review",
    "",
    `Status: ${report.delivery.status}${report.delivery.path ? `; path: ${report.delivery.path}` : ""}.`,
    ...summarizeIssues(deliveryIssues),
  ].join("\n");
  const content = report.contentView.markdown || (report.contentView.message ? `_${report.contentView.message}_` : "_No text view available._");
  const contentBudget = Math.max(80, maxChars - prefix.length - suffix.length - 4);
  const boundedContent = truncateMarkdown(content, contentBudget).markdown;
  return truncateMarkdown(`${prefix}\n${boundedContent}\n\n${suffix}`, maxChars);
}

async function materializeReviewInput(input, options, maxBytes) {
  const format = modelFormat(input);
  if (format) {
    const exported = await exportModel(input, format, options.exportOptions || {});
    if (exported.bytes.byteLength > maxBytes) throw new RangeError(`Review output has ${exported.bytes.byteLength} bytes and exceeds maxBytes (${maxBytes}).`);
    return { bytes: new Uint8Array(exported.bytes), format, path: options.outputPath ? path.resolve(options.outputPath) : undefined };
  }
  const materialized = await readBoundedBytes(input, maxBytes, "Review output");
  return { ...materialized, format: selectFormat(input, materialized.bytes, options) };
}

async function semanticReview(model, maxChars, options = {}) {
  if (!model) return { status: "failed", ok: false, issues: [reviewIssue("reimportFailed", options.importError || "The exported artifact could not be reopened.")] };
  try {
    const verification = model.verify({ ...(options.verifyOptions || {}), maxChars });
    const inspection = typeof model.inspect === "function" ? model.inspect({ ...(options.inspectOptions || {}), maxChars }) : { ndjson: "", truncated: false };
    const issues = verification.issues || [];
    return {
      status: hasHardIssue(issues) ? "failed" : issues.length ? "passed-with-warnings" : "passed",
      ok: !hasHardIssue(issues),
      issues,
      recordCounts: recordCounts(inspection.ndjson),
      inspection: { ndjson: inspection.ndjson, truncated: Boolean(inspection.truncated) },
      verification: { ndjson: verification.ndjson, truncated: Boolean(verification.truncated) },
    };
  } catch (error) {
    return { status: "failed", ok: false, issues: [reviewIssue("semanticReviewFailed", boundedMessage(error))] };
  }
}

async function structuralReview(bytes, format, maxChars, maxBytes, options = {}) {
  try {
    const inspection = await inspectPackage(bytes, format, { ...(options.inspectOptions || {}), maxBytes, maxChars });
    const issues = inspection.issues || [];
    const summary = inspection.summary || inspection.records?.[0] || {};
    const ok = inspection.ok !== false && !hasHardIssue(issues);
    return {
      status: ok ? issues.length ? "passed-with-warnings" : "passed" : "failed",
      ok,
      summary,
      issues,
      ndjson: inspection.ndjson,
      truncated: Boolean(inspection.truncated),
    };
  } catch (error) {
    return { status: "failed", ok: false, summary: {}, issues: [reviewIssue("structuralReviewFailed", boundedMessage(error))] };
  }
}

async function layoutReview(model, bytes, format, maxChars, options = {}) {
  if (options.enabled === false) return { status: "skipped", ok: true, issues: [], scope: "none" };
  if (!model) return { status: "blocked", ok: false, issues: [reviewIssue("layoutReviewBlocked", "Layout review requires a successfully reopened artifact.")], scope: "none" };
  try {
    const renderOptions = { ...(options.renderOptions || {}), maxChars };
    let target = model;
    let scope = "modeled representative render";
    if (format === "pdf") {
      const page = positiveInteger(renderOptions.page ?? 1, 1, "layout.renderOptions.page");
      target = {
        __officeKitReviewKind: "pdf",
        render: () => PdfFile.renderPdf(new FileBlob(bytes, { type: FORMAT_DETAILS.pdf.type }), {
          dpi: 144,
          format: "png",
          ...renderOptions,
          page,
        }),
      };
      scope = `native PDF page ${page}`;
    }
    const result = await visualQaApi.visualQaArtifact(target, renderOptions);
    const issues = result.issues || [];
    const ok = !hasHardIssue(issues);
    return {
      status: ok ? issues.length ? "passed-with-warnings" : "passed" : "failed",
      ok,
      scope,
      summary: result.summary,
      issues,
      ndjson: result.ndjson,
      truncated: Boolean(result.truncated),
    };
  } catch (error) {
    return { status: "failed", ok: false, issues: [reviewIssue("layoutReviewFailed", boundedMessage(error))], scope: "none" };
  }
}

async function deliveryReview(materialized, source, options, maxBytes) {
  const issues = [];
  const outputPath = materialized.path || (options.outputPath ? path.resolve(options.outputPath) : undefined);
  let sourceInfo;
  if (source != null) sourceInfo = await readBoundedBytes(source, maxBytes, "Review source");
  if (outputPath && sourceInfo?.path && await canonicalPath(outputPath) === await canonicalPath(sourceInfo.path)) {
    issues.push(reviewIssue("inputOutputCollision", "The final output path resolves to the read-only source path.", "error", { outputPath, sourcePath: sourceInfo.path }));
  }
  if (!outputPath) issues.push(reviewIssue("missingOutputPath", "No final output path was supplied; delivery publication is not yet proven.", "warning"));
  const hardFailure = hasHardIssue(issues);
  return {
    status: hardFailure ? "failed" : issues.length ? "ready-with-warnings" : "ready",
    ok: !hardFailure,
    path: outputPath,
    type: FORMAT_DETAILS[materialized.format].type,
    bytes: materialized.bytes.byteLength,
    sha256: sha256(materialized.bytes),
    sourceSha256: sourceInfo ? sha256(sourceInfo.bytes) : undefined,
    sourceChanged: sourceInfo ? sha256(sourceInfo.bytes) !== sha256(materialized.bytes) : undefined,
    issues,
  };
}

/**
 * Reopen a final Office/PDF artifact and produce one bounded post-edit review
 * report. AnyDoc supplies only the compact text view; native OfficeKit checks
 * remain authoritative for semantic, structural, render, and delivery facts.
 */
export async function reviewArtifact(input, options = {}) {
  const maxBytes = positiveInteger(options.maxBytes, DEFAULT_MAX_BYTES, "maxBytes");
  const maxContentChars = positiveInteger(options.maxContentChars, DEFAULT_MAX_CONTENT_CHARS, "maxContentChars");
  const maxInspectChars = positiveInteger(options.maxInspectChars, DEFAULT_MAX_INSPECT_CHARS, "maxInspectChars");
  const maxSummaryChars = positiveInteger(options.maxSummaryChars, DEFAULT_MAX_SUMMARY_CHARS, "maxSummaryChars");
  const visualReview = options.visualReview ?? "unavailable";
  if (!VISUAL_REVIEW_STATUSES.has(visualReview)) throw new TypeError("visualReview must be complete, unavailable, or requires-human.");
  const contentViewRequested = options.contentView === true || options.contentView === "anydoc";
  if (![undefined, false, true, "none", "anydoc"].includes(options.contentView)) {
    throw new TypeError("contentView must be anydoc, none, true, or false.");
  }

  const materialized = await materializeReviewInput(input, options, maxBytes);
  let model;
  let importError;
  const importOptions = { ...(options.importOptions || {}) };
  if (materialized.format === "pdf") importOptions.maxBytes = maxBytes;
  try { model = await importModel(materialized.bytes, materialized.format, importOptions); }
  catch (error) { importError = boundedMessage(error); }

  const semantic = await semanticReview(model, maxInspectChars, { ...options, importError });
  const structural = await structuralReview(materialized.bytes, materialized.format, maxInspectChars, maxBytes, options);
  const layout = structural.ok
    ? await layoutReview(model, materialized.bytes, materialized.format, maxInspectChars, { enabled: options.layout !== false, renderOptions: options.renderOptions })
    : { status: "blocked", ok: false, issues: [reviewIssue("layoutReviewBlocked", "Structural review failed before rendering untrusted output.")], scope: "none" };
  const contentView = !contentViewRequested
    ? { status: "not-requested", requested: false, provider: "anydoc", providerVersion: ANYDOC_VERSION, format: materialized.format }
    : structural.ok
      ? await anyDocContentView(materialized.bytes, materialized.format, { maxChars: maxContentChars })
      : { status: "blocked", requested: true, provider: "anydoc", providerVersion: ANYDOC_VERSION, format: materialized.format, reason: "structural-review-failed" };
  const delivery = await deliveryReview(materialized, options.source, options, maxBytes);

  let baseline;
  let baselineReview;
  // Source-bound work may inherit layout findings from the imported source.
  // Treat an explicit baseline as authoritative, otherwise use `source` so
  // unchanged findings are preserved as evidence rather than new failures.
  const baselineInput = options.baseline ?? options.source;
  if (baselineInput != null) {
    baselineReview = await reviewArtifact(baselineInput, {
      ...options,
      baseline: undefined,
      source: undefined,
      outputPath: undefined,
      contentView: "none",
    });
    if (baselineReview.format !== materialized.format) {
      throw new TypeError(`Review baseline format ${baselineReview.format} does not match output format ${materialized.format}.`);
    }
    baseline = applyBaselineReview({ semantic, structural, layout }, baselineReview);
  }
  const hardFailure = !semantic.ok || !structural.ok || !layout.ok || !delivery.ok;
  const limitations = !hardFailure && (
    visualReview !== "complete"
    || (contentView.requested && contentView.status !== "ready")
    || semantic.status === "passed-with-warnings"
    || structural.status === "passed-with-warnings"
    || layout.status !== "passed"
    || delivery.status !== "ready"
  );
  const report = {
    schemaVersion: 1,
    artifactKind: FORMAT_DETAILS[materialized.format].artifactKind,
    format: materialized.format,
    verdict: hardFailure ? "failed" : limitations ? "passed-with-limitations" : "passed",
    semantic,
    structural,
    layout,
    contentView,
    visualReview,
    delivery,
    ...(baselineReview ? {
      baseline: {
        sourceSha256: baselineReview.delivery.sha256,
        verdict: baselineReview.verdict,
        matchedIssues: baseline.matchedIssues,
        newIssues: baseline.newIssues,
        policy: "unchanged-issues-only",
      },
    } : {}),
  };
  report.summary = createReviewMarkdown(report, maxSummaryChars);
  return report;
}
