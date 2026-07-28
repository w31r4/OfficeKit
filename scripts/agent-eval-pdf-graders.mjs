import { spawnSync } from "node:child_process";
import { realpathSync } from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const oracleScript = path.join(scriptDirectory, "agent-eval-pdf-oracle.py");
const supportedCases = new Set([
  "pdf-bounded-contract-id-replace",
  "pdf-source-bound-text-highlight",
  "pdf-overflow-replace-refusal",
  "pdf-acroform-visible-preserved",
  "pdf-attachment-quarantine-inventory",
  "pdf-active-content-public-sanitize",
  "pdf-redact-multichannel-secret",
  "pdf-greenfield-accessible-report",
  "pdf-merge-reorder-stamp-links",
  "pdf-encrypted-owner-policy-boundary",
  "pdf-annotation-reply-resolve-boundary",
  "pdf-auto-pdfua-overclaim-refusal",
  "pdf-mixed-scan-ocr-boundary",
  "pdf-dynamic-xfa-boundary",
  "pdf-print-production-boundary",
  "pdf-richmedia-opaque-preservation",
  "pdf-docmdp-forbidden-title-edit",
  "pdf-docmdp-allowed-field-fill",
  "pdf-damaged-xref-recovery",
  "pdf-cross-page-table-extraction",
]);
const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const boundaryRefusalCases = new Map([
  ["pdf-encrypted-owner-policy-boundary", { boundary: "encrypted-owner-policy", source: "inputs/owner-policy-aes256.pdf", credential: "inputs/credentials/user-password.json" }],
  ["pdf-annotation-reply-resolve-boundary", { boundary: "annotation-reply-chain", source: "inputs/source.pdf" }],
  ["pdf-auto-pdfua-overclaim-refusal", { boundary: "pdfua-overclaim", source: "inputs/source.pdf" }],
  ["pdf-mixed-scan-ocr-boundary", { boundary: "mixed-scan-ocr", source: "inputs/source.pdf" }],
  ["pdf-dynamic-xfa-boundary", { boundary: "dynamic-xfa", source: "inputs/source.pdf" }],
  ["pdf-print-production-boundary", { boundary: "print-production", source: "inputs/source.pdf" }],
  ["pdf-richmedia-opaque-preservation", { boundary: "richmedia-opaque", source: "inputs/source.pdf" }],
  ["pdf-docmdp-forbidden-title-edit", { boundary: "docmdp-p1", source: "inputs/source.pdf" }],
]);

function check(id, category, passed, details = {}) {
  return { id, category, gate: false, passed: Boolean(passed), ...details };
}

function gate(id, category, passed, details = {}) {
  return { id, category, gate: true, passed: Boolean(passed), ...details };
}

function sameNumber(left, right, tolerance = 0.001) {
  return Number.isFinite(left) && Number.isFinite(right) && Math.abs(left - right) <= tolerance;
}

function samePageBoxes(sourcePages = [], outputPages = []) {
  return sourcePages.length === outputPages.length && sourcePages.every((source, index) => {
    const output = outputPages[index];
    return output
      && sameNumber(source.width, output.width)
      && sameNumber(source.height, output.height)
      && source.rotation === output.rotation;
  });
}

function sameBoundingBox(left = [], right = [], tolerance = 0.01) {
  return left.length === 4 && right.length === 4 && left.every((value, index) => sameNumber(value, right[index], tolerance));
}

function auditProvider(audit) {
  const provider = audit?.provider;
  if (typeof provider === "string") return provider;
  return provider?.actual || provider?.actualProvider || provider?.actual_inspection_provider || provider?.selected || provider?.name || provider?.provider || audit?.actualProvider || "";
}

function auditFallback(audit) {
  const provider = audit?.provider;
  const values = [
    provider?.silentFallback,
    provider?.silent_fallback,
    provider?.fallbackUsed,
    provider?.fallback_used,
    provider?.providerSwitched,
    provider?.provider_switched,
    audit?.silentFallback,
    audit?.silent_fallback,
    audit?.fallbackUsed,
    audit?.fallback_used,
  ].filter((value) => value !== undefined);
  return values.length ? values.every((value) => value === false || value === "false") : null;
}

function auditProviderVersion(audit) {
  const provider = audit?.provider;
  return provider?.version || provider?.actualProviderVersion || provider?.actual_inspection_provider_version || provider?.providerVersion || audit?.providerVersion || audit?.provider_version || "";
}

function auditSourceHash(audit) {
  return audit?.source?.sha256 || audit?.sourceSha256 || audit?.source_sha256 || "";
}

function auditOutputHash(audit) {
  return audit?.output?.sha256 || audit?.outputSha256 || audit?.output_sha256 || "";
}

function auditSaveRecord(audit) {
  const policy = audit?.savePolicy || audit?.save_policy || audit?.saveStrategy || audit?.save_strategy;
  return policy && typeof policy === "object" && !Array.isArray(policy) ? policy : {};
}

function auditSaveStrategy(audit) {
  const rawPolicy = audit?.savePolicy || audit?.save_policy || audit?.saveStrategy || audit?.save_strategy;
  const policy = auditSaveRecord(audit);
  return typeof rawPolicy === "string"
    ? rawPolicy
    : policy.strategy || policy.mode || policy.selected || audit?.strategy || "";
}

function auditOperationRecords(audit) {
  const operation = audit?.operation;
  if (operation && typeof operation === "object" && !Array.isArray(operation)) {
    const ordered = Array.isArray(operation.orderedOperations)
      ? operation.orderedOperations.filter((value) => value && typeof value === "object")
      : [];
    return [operation, ...ordered];
  }
  if (Array.isArray(operation)) return operation.filter((value) => value && typeof value === "object");
  if (Array.isArray(audit?.operations)) return audit.operations.filter((value) => value && typeof value === "object");
  return [];
}

function auditOperation(audit) {
  const operation = audit?.operation;
  if (typeof operation === "string") return operation;
  const records = auditOperationRecords(audit);
  return records.map((value) => value.type || value.operation || value.name || value.performed || "").join(" ");
}

function auditOperationRecord(audit) {
  const records = auditOperationRecords(audit);
  return records.find((record) => record.mode === "mutation" || /(?:permission|edit|replace|redact|fill|flatten|sign|mutation|refus)/i.test(String(record.type || record.operation || record.name || ""))) || records[0] || null;
}

function auditRefusalOperationUnexecuted(audit) {
  return auditOperationRecords(audit).some((operation) => Boolean(operation.type || operation.operation || operation.name) && (
    operation.executed === false
    || operation.performed === false
    || operation.performed === "none"
    || operation.mutationAttempted === false
    || /(?:refused|not.?attempted|not.?performed|no.?mutation)/i.test(String(operation.result || ""))
  ));
}

function auditRefusalNoArtifactClaim(audit) {
  const validation = audit?.validation || {};
  const artifactChecks = validation.artifactChecks || validation.artifact_checks || validation.artifactCheck || validation.artifact_check || {};
  const outputDirectoryPolicy = validation.outputDirectoryPolicy || validation.outputsDirectory || {};
  const saveRecord = auditSaveRecord(audit);
  const verification = audit?.verification || {};
  const validationChecks = Array.isArray(validation.checks) ? validation.checks : [];
  const auditOnly = saveRecord.allowed_output === "audit_only"
    || saveRecord.selected === "none" && /refused/i.test(String(saveRecord.publication || ""))
    || /(?:failed.closed|audit.only)/i.test(String(saveRecord.publication || ""));
  const noModifiedArtifact = (
    audit?.delivered_modified_pdf === false
    || validation.no_partial_modified_pdf_in_outputs === true
    || audit?.noPartialOutput === true
    || artifactChecks.modifiedPdfCreated === false
    || artifactChecks.modifiedPdfProduced === false
    || artifactChecks.modifiedPdfPresent === false
    || artifactChecks.outputIsNull === true
    || artifactChecks.modifiedArtifactCreated === false
    || artifactChecks.modifiedArtifactProduced === false
    || verification.modifiedArtifactPresent === false
    || saveRecord.artifactWritten === false
    || saveRecord.artifact_written === false
    || validationChecks.some((entry) => /(?:modified|artifact).*absent/i.test(String(entry?.check || "")) && /^(?:passed|true)$/i.test(String(entry?.result ?? entry?.passed)))
    || outputDirectoryPolicy.modifiedPdfPresent === false && outputDirectoryPolicy.partialArtifactPresent === false
  );
  const noPartialArtifact = (
    validation.no_partial_modified_pdf_in_outputs === true
    || audit?.noPartialOutput === true
    || artifactChecks.partialArtifactPresent === false
    || artifactChecks.partialArtifactProduced === false
    || saveRecord.partialResultsWritten === false
    || saveRecord.partial_results_written === false
    || outputDirectoryPolicy.partialArtifactPresent === false
    || auditOnly
  );
  const outputIsNullOrOmitted = !Object.hasOwn(audit || {}, "output") || audit.output === null;
  return outputIsNullOrOmitted && audit?.delivered_modified_pdf !== true && noModifiedArtifact && noPartialArtifact;
}

function auditRefusalSavePolicy(audit) {
  const saveRecord = auditSaveRecord(audit);
  const noMutation = saveRecord.executed === false
    || saveRecord.performed === false
    || saveRecord.attempted === false
    || saveRecord.sourceOverwrite === false
    || saveRecord.modifiedArtifactAllowed === false
    || saveRecord.artifactWritten === false
    || saveRecord.artifact_written === false
    || saveRecord.mode === "failed_closed"
    || saveRecord.selected === "none"
    || auditRefusalOperationUnexecuted(audit)
    || /(?:failed.closed|audit.only|refused.*mutation)/i.test(String(saveRecord.publication || saveRecord.decision || ""));
  const sourcePreserved = saveRecord.sourcePreserved === true
    || saveRecord.sourceOverwrite === false
    || audit?.source?.preserved_unchanged === true
    || audit?.source?.preserved === true
    || audit?.validation?.sourceImmutable === true
    || audit?.validation?.sourceIdentity?.sourcePreserved === true
    || audit?.validation?.sourceIntegrity?.sourceOverwritten === false;
  return Boolean(auditSaveStrategy(audit)) && noMutation && sourcePreserved && auditRefusalNoArtifactClaim(audit);
}

export function extractCompletedCommands(trace) {
  const commands = new Map();
  for (const line of String(trace || "").split(/\r?\n/)) {
    if (!line.trim()) continue;
    let event;
    try { event = JSON.parse(line); } catch { continue; }
    const item = event?.item;
    if (event?.type === "item.completed" && item?.type === "command_execution" && typeof item.command === "string") {
      commands.set(item.id || `${commands.size}`, item.command);
    }
  }
  return [...commands.values()];
}

function boundedTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const probeIndex = commandText.search(/pymupdf_edit\.py\s+probe\b/i);
  const editIndex = commandText.search(/pymupdf_edit\.py\s+edit\b/i);
  const bypassPatterns = [
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\bxref_set_(?:key|stream)\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /pymupdf/i.test(String(auditProvider(audit))), { expected: "pymupdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^sanitize$/i.test(String(auditSaveStrategy(audit))), { expected: "sanitize", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:capability-probe", "trace", probeIndex >= 0 && editIndex >= 0 && probeIndex < editIndex, { expected: "shipped PyMuPDF capability probe before mutation", actual: { probeObserved: probeIndex >= 0, editObserved: editIndex >= 0, preflightOrder: probeIndex >= 0 && editIndex >= 0 && probeIndex < editIndex } }),
    check("pdf-trace:typed-edit-primitive", "trace", /pymupdf_edit\.py\s+edit\b/i.test(commandText) && /replace[_ -]?text/i.test(operation), { expected: "pymupdf_edit.py edit + replace_text", actual: operation || "unreported" }),
    check("pdf-trace:no-content-stream-bypass", "trace", !bypassPatterns.some((pattern) => pattern.test(commandText)), { forbidden: bypassPatterns.map(String) }),
  ];
}

function overflowTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const providerProbe = /pymupdf_edit\.py\s+probe\b/i.test(commandText)
    || /\b(?:import|from)\s+(?:fitz|pymupdf)\b/i.test(commandText)
    || /importlib\.metadata[\s\S]{0,200}pymupdf/i.test(commandText);
  const mutationPatterns = [
    /pymupdf_edit\.py\s+edit\b/i,
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\badd_redact_annot\s*\(/i,
    /\bapply_redactions\s*\(/i,
    /\binsert_text(?:box)?\s*\(/i,
    /\bdoc(?:ument)?\.save\s*\(/i,
    /\bwriter\.write\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /pymupdf/i.test(String(auditProvider(audit))), { expected: "pymupdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:capability-probe", "trace", providerProbe, { expected: "PyMuPDF capability/version evidence" }),
    gate("pdf-trace:no-mutation-after-failed-preflight", "trace", !mutationPatterns.some((pattern) => pattern.test(commandText)), { forbidden: mutationPatterns.map(String) }),
  ];
}

function docMdpAuditValidation(audit) {
  const verification = audit?.validation?.signatureVerification;
  const policy = audit?.signaturePolicy;
  const docMdp = policy?.docMDP;
  return verification?.conclusion === "valid-under-selected-policy"
    && verification?.cryptographicallyValid === true
    && verification?.intact === true
    && verification?.trusted === true
    && verification?.trustPolicy === "explicit-roots"
    && typeof verification?.trustRoot?.sha256 === "string"
    && /^[a-f0-9]{64}$/i.test(verification.trustRoot.sha256)
    && verification?.bottomLine === true
    && verification?.coverage === "entire-file"
    && verification?.docMDPCompliantBeforeRequestedChange === true
    && policy?.certificationSignaturePresent === true
    && docMdp?.permissionCode === 1
    && docMdp?.requestedChangeAllowed === false
    && policy?.policyDecision === "refuse_without_mutation";
}

function boundaryRefusalTraceChecks(audit, commands, item) {
  const commandText = commands.join("\n");
  const saveRecord = auditSaveRecord(audit);
  const hasInspection = /(?:pdf_provider\.py\s+(?:check|plan)|qpdf_provider\.py\s+(?:inspect|probe)|pikepdf_provider\.py\s+(?:inspect|probe)|pyhanko(?:_sign)?_provider\.py\s+(?:inspect|verify|probe)|pymupdf_edit\.py\s+probe|mupdf\.mjs["']?\s+(?:probe|inspect)\b|\bqpdf\b|\bpdfinfo\b|\bpypdf\b|PdfFile\.(?:inspect|open))/i.test(commandText);
  const hasDocMdpTrustValidation = commands.some((command) => {
    const text = String(command);
    return /pyhanko_provider\.py["']?\s+verify\b/i.test(text)
      && /--trust-policy\s+explicit-roots\b/i.test(text)
      && /--trust-root\s+/i.test(text)
      && /--require-signature\b/i.test(text)
      && /--require-all-integrity-valid\b/i.test(text)
      && /--require-all-trusted\b/i.test(text)
      && /--require-docmdp-compliant\b/i.test(text)
      && /--require-all-bottom-line\b/i.test(text);
  });
  const hasVerifiedDocMdpTrustEvidence = hasDocMdpTrustValidation && docMdpAuditValidation(audit);
  const mutationPatterns = [
    /pymupdf_edit\.py\s+edit\b/i,
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\badd_redact_annot\s*\(/i,
    /\bapply_redactions\s*\(/i,
    /pyhanko_sign_provider\.py\s+sign\b/i,
    /\bwriter\.write\s*\(/i,
    /\bdoc(?:ument)?\.save\s*\(/i,
  ];
  return [
    check("pdf-trace:provider-declared", "trace", Boolean(String(auditProvider(audit)).trim()), { actual: auditProvider(audit) || "unreported" }),
    check("pdf-trace:refusal-save-policy", "trace", auditRefusalSavePolicy(audit), { expected: "a declared no-mutation policy with immutable source and no published artifact", actual: saveRecord }),
    check("pdf-trace:source-inspection", "trace", hasInspection, { expected: "an explicit PDF/provider inspection before refusal" }),
    ...(item.id === "pdf-docmdp-forbidden-title-edit"
      ? [gate("pdf-security:docmdp-explicit-root-validation", "security", hasVerifiedDocMdpTrustEvidence, {
        expected: "pyHanko verification plus audit-bound explicit-root integrity, trust, DocMDP, and bottom-line evidence",
        actual: { command: hasDocMdpTrustValidation, audit: docMdpAuditValidation(audit) },
      })]
      : []),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    gate("pdf-trace:no-mutation-command", "trace", !mutationPatterns.some((pattern) => pattern.test(commandText)), { forbidden: mutationPatterns.map(String) }),
  ];
}

function activeContentTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const invocationIndex = (pattern) => commands.findIndex((command) => {
    const match = String(command).match(pattern);
    if (!match) return false;
    const tail = String(command).slice((match.index || 0) + match[0].length);
    return !/^\s+(?:--help|-h)\b/i.test(tail);
  });
  const probeIndex = invocationIndex(/pymupdf_edit\.py\s+probe\b/i);
  const planIndex = invocationIndex(/pdf_provider\.py\s+plan\b/i);
  const editIndex = invocationIndex(/pymupdf_edit\.py\s+edit\b/i);
  const afterEdit = editIndex >= 0 ? commands.slice(editIndex).join("\n") : "";
  const residueAfterEdit = /residue_scan\.py\b/i.test(afterEdit);
  const renderAfterEdit = /\bpdftoppm\b/i.test(afterEdit);
  const auditAfterEdit = /pdf_audit\.py\s+validate\b/i.test(afterEdit);
  const bypassPatterns = [
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\bxref_set_(?:key|stream)\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /pymupdf/i.test(String(auditProvider(audit))), { expected: "pymupdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^sanitize$/i.test(String(auditSaveStrategy(audit))), { expected: "sanitize", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:preflight-audit", "trace", audit?.preflight?.probeCompleted === true && audit?.preflight?.planCompleted === true, { actual: audit?.preflight || "unreported" }),
    check("pdf-trace:probe-plan-before-mutation", "trace", probeIndex >= 0 && planIndex >= 0 && editIndex > probeIndex && editIndex > planIndex, {
      actual: {
        probeObserved: probeIndex >= 0,
        planObserved: planIndex >= 0,
        editObserved: editIndex >= 0,
        probeCommandIndex: probeIndex,
        planCommandIndex: planIndex,
        editCommandIndex: editIndex,
      },
    }),
    check("pdf-trace:typed-scrub-primitive", "trace", /pymupdf_edit\.py\s+edit\b/i.test(commandText) && /scrub|active[_ -]?content/i.test(operation), { actual: operation || "unreported" }),
    check("pdf-trace:post-mutation-residue-scan", "trace", residueAfterEdit, { actual: { editObserved: editIndex >= 0, postMutationResidueScanObserved: residueAfterEdit } }),
    check("pdf-trace:post-mutation-poppler-render", "trace", renderAfterEdit, { actual: { editObserved: editIndex >= 0, postMutationRenderObserved: renderAfterEdit } }),
    check("pdf-trace:audit-byte-validation", "trace", auditAfterEdit, { actual: { postMutationAuditValidationObserved: auditAfterEdit } }),
    gate("pdf-trace:no-content-stream-bypass", "trace", !bypassPatterns.some((pattern) => pattern.test(commandText)), { forbidden: bypassPatterns.map(String) }),
  ];
}

function multichannelRedactionTraceChecks(audit, commands) {
  const operations = auditOperationRecords(audit);
  const ocrIndex = operations.findIndex((operation) => operation.type === "redact_ocr_text");
  const textIndex = operations.findIndex((operation) => operation.type === "redact_text");
  const scrubIndex = operations.findIndex((operation) => operation.type === "scrub" || operation.type === "active_content_cleanup");
  const commandText = commands.join("\n");
  const editIndex = commands.findIndex((command) => /pymupdf_edit\.py\s+edit\b/i.test(command));
  const postEditCommands = editIndex >= 0 ? commands.slice(editIndex).join("\n") : "";
  return [
    ...activeContentTraceChecks(audit, commands),
    check("pdf-trace:typed-selectable-and-raster-redaction", "trace", ocrIndex >= 0 && textIndex >= 0 && scrubIndex >= 0 && ocrIndex < textIndex, {
      expected: "redact_ocr_text before redact_text, followed by scrub/cleanup",
      actual: operations.map((operation) => operation.type || operation.operation || "unreported"),
    }),
    check("pdf-trace:ocr-probe-required", "trace", /pymupdf_edit\.py\s+probe\b[^\n]*--require-ocr/i.test(commandText), { expected: "pymupdf OCR capability probe" }),
    check("pdf-trace:strict-post-mutation-residue-scan", "trace", /residue_scan\.py\b/i.test(postEditCommands) && /--require-ocr\b/i.test(postEditCommands) && /--require-single-revision\b/i.test(postEditCommands), {
      expected: "post-mutation residue scan with OCR and revision requirements",
    }),
  ];
}

function completedInvocation(commands, pattern) {
  for (let commandIndex = 0; commandIndex < commands.length; commandIndex += 1) {
    const command = String(commands[commandIndex]);
    const expression = new RegExp(pattern.source, pattern.flags.replace("g", "") + "g");
    for (const match of command.matchAll(expression)) {
      const tail = command.slice((match.index || 0) + match[0].length);
      if (/^\s+(?:--help|-h)\b/i.test(tail)) continue;
      return { commandIndex, offset: match.index || 0 };
    }
  }
  return null;
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function qpdfProviderPath(value) {
  return /(?:^|[/\\'"\s])qpdf_provider\.py\b/i.test(String(value || ""));
}

function qpdfProviderBindings(command) {
  const bindings = new Set();
  const assignment = /(?:^|[\s;])(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)=(?:"([^"]*)"|'([^']*)'|([^\s;]+))/g;
  for (const match of String(command || "").matchAll(assignment)) {
    const [, name, doubleQuoted, singleQuoted, unquoted] = match;
    if (qpdfProviderPath(doubleQuoted ?? singleQuoted ?? unquoted ?? "")) bindings.add(name);
  }
  return bindings;
}

function qpdfProviderInvocationPattern(action, bindings = []) {
  const variables = [...new Set(bindings)].map(escapeRegExp);
  const variableInvocation = variables.length > 0
    ? `|["']?\\$\\{?(?:${variables.join("|")})\\}?["']?`
    : "";
  const directInvocation = String.raw`(?:^|[\s/'"])qpdf_provider\.py["']?`;
  // A direct adapter path is always allowed. A variable invocation is allowed
  // only when its name came from qpdfProviderBindings() for this shell command
  // or its containing rewrite command.
  return new RegExp(`(?:${directInvocation}${variableInvocation})\\s+${escapeRegExp(action)}\\b`, "i");
}

function qpdfProviderInvocationMatches(command, action, bindings = []) {
  const pattern = qpdfProviderInvocationPattern(action, bindings);
  return [...String(command || "").matchAll(new RegExp(pattern.source, `${pattern.flags}g`))];
}

function completedQpdfProviderInvocation(commands, action) {
  for (let commandIndex = 0; commandIndex < commands.length; commandIndex += 1) {
    const command = String(commands[commandIndex]);
    const bindings = qpdfProviderBindings(command);
    for (const match of qpdfProviderInvocationMatches(command, action, bindings)) {
      const tail = command.slice((match.index || 0) + match[0].length);
      if (/^\s+(?:--help|-h)\b/i.test(tail)) continue;
      return { commandIndex, offset: match.index || 0, bindings: [...bindings] };
    }
  }
  return null;
}

function hasQpdfProviderInvocation(command, action, inheritedBindings = []) {
  const text = String(command || "");
  const bindings = new Set([
    ...qpdfProviderBindings(text),
    ...(Array.isArray(inheritedBindings) ? inheritedBindings : []),
  ]);
  return qpdfProviderInvocationMatches(text, action, bindings).length > 0;
}

function invocationBefore(left, right) {
  return Boolean(left && right && (
    left.commandIndex < right.commandIndex
    || left.commandIndex === right.commandIndex && left.offset < right.offset
  ));
}

function commandTextAfter(commands, position) {
  if (!position) return "";
  return [String(commands[position.commandIndex]).slice(position.offset), ...commands.slice(position.commandIndex + 1)].join("\n");
}

function highlightColorMatches(actual, expected, tolerance = 0.001) {
  return Array.isArray(actual)
    && Array.isArray(expected)
    && actual.length === 3
    && expected.length === 3
    && actual.every((value, index) => sameNumber(Number(value), Number(expected[index]), tolerance));
}

function highlightQuadMatchesSourceText(highlight, sourceTarget, pageHeight, tolerance = 4) {
  const values = highlight?.quadPoints;
  const bbox = sourceTarget?.bbox;
  if (!Array.isArray(values) || values.length < 8 || values.length % 8 !== 0 || !Array.isArray(bbox) || bbox.length !== 4 || !Number.isFinite(pageHeight)) return false;
  const [x0, top, x1, bottom] = bbox.map(Number);
  const minY = pageHeight - bottom - tolerance;
  const maxY = pageHeight - top + tolerance;
  return values.every((value, index) => {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) return false;
    return index % 2 === 0
      ? numeric >= x0 - tolerance && numeric <= x1 + tolerance
      : numeric >= minY && numeric <= maxY;
  });
}

function sourceBoundHighlightTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const operationRecord = auditOperationRecord(audit);
  const probe = completedInvocation(commands, /mupdf\.mjs["']?\s+probe\b/i);
  const inspect = completedInvocation(commands, /mupdf\.mjs["']?\s+inspect\b/i);
  const edit = completedInvocation(commands, /mupdf\.mjs["']?\s+edit\b/i);
  const afterEdit = commandTextAfter(commands, edit);
  const outputInspect = /mupdf\.mjs["']?\s+inspect\b/i.test(afterEdit);
  const outputRender = /mupdf\.mjs["']?\s+render\b/i.test(afterEdit);
  const directMutationPatterns = [
    /\baddAnnotation\s*\(/i,
    /\bsetQuadPoints\s*\(/i,
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\bxref_set_(?:key|stream)\s*\(/i,
  ];
  const callerGeometryFields = ["bbox", "rect", "quadPoints", "quads", "point"];
  return [
    check("pdf-trace:provider", "trace", /^mupdf$/i.test(String(auditProvider(audit))), { expected: "mupdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^rewrite$/i.test(String(auditSaveStrategy(audit))), { expected: "rewrite", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:probe-inspect-before-edit", "trace", invocationBefore(probe, edit) && invocationBefore(inspect, edit), { actual: { probe: probe?.commandIndex ?? -1, inspect: inspect?.commandIndex ?? -1, edit: edit?.commandIndex ?? -1 } }),
    check("pdf-trace:typed-source-bound-highlight", "trace", Boolean(edit) && operation === "add_text_highlight" && operationRecord?.type === "add_text_highlight", { actual: operation || "unreported" }),
    check("pdf-trace:reinspect-and-render-output", "trace", outputInspect && outputRender, { actual: { outputInspect, outputRender } }),
    gate("pdf-trace:no-caller-coordinate-or-direct-mutation", "trace", !callerGeometryFields.some((field) => Object.hasOwn(operationRecord || {}, field)) && !directMutationPatterns.some((pattern) => pattern.test(commandText)), { forbidden: callerGeometryFields.concat(directMutationPatterns.map(String)) }),
  ];
}

function acroformTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const inspect = completedInvocation(commands, /pypdf_edit\.py\s+inspect\b/i);
  const checkProvider = completedInvocation(commands, /pdf_provider\.py\s+check\b[^\n]*--provider\s+pypdf\b/i);
  const plan = completedInvocation(commands, /pdf_provider\.py\s+plan\b/i);
  const fill = completedInvocation(commands, /pypdf_edit\.py\s+fill-form\b/i);
  const afterFill = commandTextAfter(commands, fill);
  const bypassPatterns = [
    /\bPdfWriter\s*\(/,
    /\bupdate_page_form_field_values\s*\(/,
    /\bclone_document_from_reader\s*\(/,
  ];
  const positions = Object.fromEntries(Object.entries({ inspect, checkProvider, plan, fill }).map(([name, value]) => [name, value?.commandIndex ?? -1]));
  return [
    check("pdf-trace:provider", "trace", /pypdf/i.test(String(auditProvider(audit))), { expected: "pypdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^incremental$/i.test(String(auditSaveStrategy(audit))), { expected: "incremental", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:preflight-audit", "trace", audit?.preflight?.probeCompleted === true && audit?.preflight?.planCompleted === true, { actual: audit?.preflight || "unreported" }),
    check("pdf-trace:inspect-check-plan-before-mutation", "trace", invocationBefore(inspect, fill) && invocationBefore(checkProvider, fill) && invocationBefore(plan, fill), { actual: positions }),
    check("pdf-trace:typed-fill-form-primitive", "trace", Boolean(fill) && /fill[-_ ]?form/i.test(operation), { actual: operation || "unreported" }),
    check("pdf-trace:post-mutation-poppler-render", "trace", /\bpdftoppm\b/i.test(afterFill), { actual: { fillObserved: Boolean(fill), postMutationRenderObserved: /\bpdftoppm\b/i.test(afterFill) } }),
    check("pdf-trace:audit-byte-validation", "trace", /pdf_audit\.py\s+validate\b/i.test(afterFill), { actual: { postMutationAuditValidationObserved: /pdf_audit\.py\s+validate\b/i.test(afterFill) } }),
    gate("pdf-trace:no-ad-hoc-pypdf-writer", "trace", !bypassPatterns.some((pattern) => pattern.test(commandText)), { forbidden: bypassPatterns.map(String) }),
  ];
}

function certifiedDocMdpP2TraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const probe = completedInvocation(commands, /pyhanko_certified_form_fill\.py["']?\s+probe\b/i);
  const fill = completedInvocation(commands, /pyhanko_certified_form_fill\.py["']?\s+fill\b/i);
  const afterFill = commandTextAfter(commands, fill);
  const postflightVerification = commands.some((command, index) => index > (fill?.commandIndex ?? Number.MAX_SAFE_INTEGER)
    && /pyhanko_provider\.py["']?\s+verify\b/i.test(String(command))
    && /--trust-policy\s+explicit-roots\b/i.test(String(command))
    && /--trust-root\s+/i.test(String(command))
    && /--require-signature\b/i.test(String(command))
    && /--require-all-integrity-valid\b/i.test(String(command))
    && /--require-all-trusted\b/i.test(String(command))
    && /--require-docmdp-compliant\b/i.test(String(command))
    && /--require-all-bottom-line\b/i.test(String(command)));
  const renderAfterFill = /(?:mupdf\.mjs["']?\s+render\b|\bpdftoppm\b)/i.test(afterFill);
  const bypassPatterns = [
    /\bPdfWriter\s*\(/,
    /\bupdate_page_form_field_values\s*\(/,
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\bwriter\.write\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /^pyhanko$/i.test(String(auditProvider(audit))), { expected: "pyhanko", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^incremental$/i.test(String(auditSaveStrategy(audit))), { expected: "incremental", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:controlled-probe-before-fill", "trace", invocationBefore(probe, fill), { actual: { probe: probe?.commandIndex ?? -1, fill: fill?.commandIndex ?? -1 } }),
    check("pdf-trace:typed-certified-form-primitive", "trace", Boolean(fill) && audit?.operation === "fill-certified-docmdp-p2-text-field", { actual: audit?.operation || "unreported" }),
    check("pdf-trace:postflight-explicit-root-validation", "trace", postflightVerification, { expected: "post-fill explicit-root pyHanko integrity, trust, DocMDP, and bottom-line validation" }),
    check("pdf-trace:postflight-render", "trace", renderAfterFill, { actual: { fillObserved: Boolean(fill), renderObserved: renderAfterFill } }),
    gate("pdf-trace:no-ad-hoc-form-writer", "trace", !bypassPatterns.some((pattern) => pattern.test(commandText)), { forbidden: bypassPatterns.map(String) }),
  ];
}

function attachmentTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  // Shell-safe invocations commonly quote a variable-expanded script path,
  // for example `"$SCRIPTS/pypdf_edit.py" inspect`. Accept that closing quote
  // without weakening the requirement that the shipped typed primitive ran.
  const inspect = completedInvocation(commands, /pypdf_edit\.py["']?\s+inspect\b/i);
  const checkProvider = completedInvocation(commands, /pdf_provider\.py["']?\s+check\b[^\n]*--provider\s+pypdf\b/i);
  const plan = completedInvocation(commands, /pdf_provider\.py["']?\s+plan\b[^\n]*--task\s+extract-attachments\b/i);
  const extract = completedInvocation(commands, /pypdf_edit\.py["']?\s+extract-attachments\b/i);
  const afterExtract = commandTextAfter(commands, extract);
  const auditValidatedAfterExtraction = /pdf_audit\.py["']?\s+validate\b/i.test(afterExtract);
  const manualPatterns = [
    /\bPdfReader\s*\(/,
    /\.attachment_list\b/,
    /\.attachments\b/,
    /\[\s*["']\/EmbeddedFiles["']\s*\]/,
    /\bget_data\s*\(/,
  ];
  const payloadOpenPatterns = [
    /\b(?:cat|less|more|head|tail|strings|unzip|7z|tar)\b[^\n]*outputs\/quarantine/i,
    // Codex command traces are normally wrapped in `/bin/zsh -lc`; that
    // launcher alone is not evidence that a quarantined payload was executed.
    // Match only an actual payload interpreter command at the beginning of a
    // command or after a shell separator.
    /(?:^|[;&|]\s*)(?:"?\$PYTHON_BIN"?|(?:\/[\w./-]+\/)?(?:python\d*|node))\s+[^\n]*outputs\/quarantine\//im,
    /(?:^|[;&|]\s*)(?:\/[\w./-]+\/)?(?:bash|sh|zsh)\s+(?!-lc\b)[^\n]*outputs\/quarantine\//im,
    /\bchmod\b[^\n]*outputs\/quarantine\//i,
  ];
  const positions = Object.fromEntries(Object.entries({ inspect, checkProvider, plan, extract }).map(([name, value]) => [name, value?.commandIndex ?? -1]));
  return [
    check("pdf-trace:provider", "trace", /pypdf/i.test(String(auditProvider(audit))), { expected: "pypdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^read-only$/i.test(String(auditSaveStrategy(audit))), { expected: "read-only", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:preflight-audit", "trace", audit?.preflight?.probeCompleted === true && audit?.preflight?.planCompleted === true, { actual: audit?.preflight || "unreported" }),
    check("pdf-trace:inspect-check-plan-before-extraction", "trace", invocationBefore(inspect, extract) && invocationBefore(checkProvider, extract) && invocationBefore(plan, extract), { actual: positions }),
    check("pdf-trace:typed-attachment-primitive", "trace", Boolean(extract) && /extract[-_ ]?attachments/i.test(auditOperation(audit)), { actual: auditOperation(audit) || "unreported" }),
    check("pdf-trace:audit-byte-validation", "trace", auditValidatedAfterExtraction, { actual: { extractObserved: Boolean(extract), postExtractionAuditValidationObserved: auditValidatedAfterExtraction } }),
    gate("pdf-trace:no-ad-hoc-pypdf-extraction", "trace", !manualPatterns.some((pattern) => pattern.test(commandText)), { forbidden: manualPatterns.map(String) }),
    gate("pdf-trace:no-payload-open-or-execution", "trace", !payloadOpenPatterns.some((pattern) => pattern.test(commandText)), { forbidden: payloadOpenPatterns.map(String) }),
  ];
}

function accessibleReportTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const manualWriterPatterns = [
    /%PDF-\d/i,
    /\bPdfWriter\s*\(/,
    /\breportlab\b/i,
    /\bpymupdf\b|\bfitz\b/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /^artifact-tool$/i.test(String(auditProvider(audit))), { expected: "artifact-tool", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^rewrite$/i.test(String(auditSaveStrategy(audit))), { expected: "rewrite", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:typed-greenfield-primitive", "trace", /accessible-board-report\.mjs/i.test(commandText) && /create[-_ ]?accessible[-_ ]?report/i.test(operation), { expected: "shipped accessible-board-report.mjs + create-accessible-report", actual: { operation, exampleObserved: /accessible-board-report\.mjs/i.test(commandText) } }),
    check("pdf-trace:poppler-review", "trace", audit?.validation?.poppler?.status === "passed" && audit?.validation?.poppler?.pages?.length === 6, { actual: audit?.validation?.poppler || "unreported" }),
    gate("pdf-trace:no-ad-hoc-pdf-writer", "trace", !manualWriterPatterns.some((pattern) => pattern.test(commandText)), { forbidden: manualWriterPatterns.map(String) }),
  ];
}

function pdfUaOverclaim(text) {
  return /PDF\/?UA\s+(?:certified|compliant|conformant|validated|passed)\b|(?<!不)(?:已通过|符合)\s*PDF\/?UA|PDF\/?UA\s*(?:已认证|已合规|认证通过)/i.test(String(text || ""));
}

export function gradeAccessibleReportEvidence({ evidence, audit, commands, finalMessage, item }) {
  const structure = evidence.structure || {};
  const output = evidence.output || {};
  const visualPages = evidence.visual?.pages || [];
  const table = structure.tables?.find((candidate) => candidate.pages?.length >= 2);
  const validation = audit?.validation || {};
  const headingLevels = item.grade.machine.headingLevels || [1, 2, 3];
  const claimText = `${JSON.stringify(audit || {})}\n${finalMessage || ""}`;
  const link = structure.links?.[0];
  const pageText = structure.pageText || [];
  return [
    check("pdf-machine:page-count", "machine", output.pageCount === item.grade.machine.pageCount, { expected: item.grade.machine.pageCount, actual: output.pageCount }),
    check("pdf-machine:catalog-title-language-tagging", "machine", structure.tagged === true && structure.language === item.grade.machine.language && Boolean(String(structure.title || "").trim()), { expected: { tagged: true, language: item.grade.machine.language, title: "non-empty" }, actual: { tagged: structure.tagged, language: structure.language, title: structure.title } }),
    check("pdf-machine:h1-h3-structure", "machine", headingLevels.every((level) => Number(structure.roles?.[`H${level}`] || 0) > 0), { expected: headingLevels, actual: structure.roles }),
    check("pdf-machine:cross-page-semantic-table", "machine", Boolean(table && table.pages.length >= 2 && table.rows >= 2 && table.headers >= 1 && table.dataCells >= 1), { actual: structure.tables || [] }),
    check("pdf-machine:figure-alt", "machine", Number(structure.roles?.Figure || 0) >= 1 && structure.figuresWithAlt === Number(structure.roles?.Figure || 0), { actual: { figures: structure.roles?.Figure || 0, withAlt: structure.figuresWithAlt } }),
    check("pdf-machine:meaningful-tagged-link", "machine", Number(structure.roles?.Link || 0) >= 1 && structure.linkObjrAssociations >= 1 && /^https?:\/\//i.test(String(link?.uri || "")) && Number.isInteger(link?.structParent), { actual: { roles: structure.roles?.Link || 0, objr: structure.linkObjrAssociations, links: structure.links } }),
    check("pdf-machine:reading-order-ids", "machine", structure.rootIds?.length >= 12 && new Set(structure.rootIds).size === structure.rootIds.length && structure.rootIds.includes("risk-register"), { actual: structure.rootIds || [] }),
    check("pdf-machine:running-artifacts", "machine", structure.artifactMarkers >= item.grade.machine.pageCount * 2, { expectedAtLeast: item.grade.machine.pageCount * 2, actual: structure.artifactMarkers }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), { actual: audit?.status || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual?.pageCount === item.grade.machine.pageCount && visualPages.length === item.grade.machine.pageCount && visualPages.every((page) => page.nonBlank && page.bytes > 1_000), { renderer: evidence.visual?.renderer, pages: visualPages }),
    check("pdf-visual:no-edge-clipping", "visual", visualPages.length > 0 && visualPages.every((page) => page.touchesEdge === false), { actual: visualPages.map(({ page, inkBBox, touchesEdge }) => ({ page, inkBBox, touchesEdge })) }),
    check("pdf-visual:consistent-page-geometry", "visual", visualPages.length > 0 && visualPages.every((page) => page.width === visualPages[0].width && page.height === visualPages[0].height), { actual: visualPages.map(({ page, width, height }) => ({ page, width, height })) }),
    check("pdf-visual:cross-page-table-readable", "visual", Boolean(table?.pages?.every((pageNumber) => /风险/.test(String(pageText[pageNumber - 1] || "")))), { tablePages: table?.pages || [], pageText: table?.pages?.map((pageNumber) => String(pageText[pageNumber - 1] || "").slice(0, 240)) || [] }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === evidence.source?.sha256 && auditOutputHash(audit) === output.sha256, { expected: { source: evidence.source?.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    gate("pdf-security:no-pdfua-overclaim", "security", !pdfUaOverclaim(claimText), { actual: pdfUaOverclaim(claimText) ? "positive PDF/UA certification claim detected" : "no overclaim detected" }),
    check("pdf-security:modeled-scope-explicit", "security", validation.modeledVerify?.status === "passed" && /modeled|PdfArtifact/i.test(String(validation.modeledVerify?.scope || "")), { actual: validation.modeledVerify || "unreported" }),
    check("pdf-security:verapdf-machine-layer-separate", "security", Boolean(validation.veraPdfMachine) && ["not-run", "completed", "completed-with-findings", "probe-failed"].includes(validation.veraPdfMachine.status) && /machine|not.*certification|No veraPDF/i.test(JSON.stringify(validation.veraPdfMachine)), { actual: validation.veraPdfMachine || "unreported" }),
    check("pdf-security:human-pdfua-required", "security", validation.humanPdfUa?.status === "required" && /No complete PDF\/UA certification|不.*完整.*PDF\/UA|人工/i.test(JSON.stringify(validation.humanPdfUa)), { actual: validation.humanPdfUa || "unreported" }),
    ...accessibleReportTraceChecks(audit, commands),
  ];
}

function mergeStampTraceChecks(audit, commands) {
  const operation = auditOperation(audit);
  const checkProvider = completedInvocation(commands, /pdf_provider\.py["']?\s+check\b/i);
  const plan = completedInvocation(commands, /pdf_provider\.py["']?\s+plan\b/i);
  const merge = completedInvocation(commands, /pypdf_edit\.py["']?\s+merge-stamp\b/i);
  const afterMerge = commandTextAfter(commands, merge);
  const popplerAfterMerge = /poppler_compare\.py["']?\s+merge-stamp\b/i.test(afterMerge);
  const auditAfterMerge = /pdf_audit\.py["']?\s+validate\b/i.test(afterMerge) && (afterMerge.match(/(?:^|\s)--input(?:=|\s)/g) || []).length >= 3;
  const manualWriterPatterns = [
    /\bPdfWriter\s*\(/,
    /\bmerge_(?:page|transformed_page)\s*\(/,
    /\bcanvas\.Canvas\s*\(/,
    /\b(?:fitz|pymupdf)\.open\s*\(/i,
    /\/Pages\b[\s\S]{0,80}\/Kids\b/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /^pypdf$/i.test(String(auditProvider(audit))), { expected: "pypdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^rewrite$/i.test(String(auditSaveStrategy(audit))), { expected: "rewrite", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:check-plan-before-mutation", "trace", invocationBefore(checkProvider, merge) && invocationBefore(plan, merge), { actual: { checkProvider: checkProvider?.commandIndex ?? -1, plan: plan?.commandIndex ?? -1, merge: merge?.commandIndex ?? -1 } }),
    check("pdf-trace:typed-merge-stamp-primitive", "trace", Boolean(merge) && /merge[-_ ]?stamp/i.test(operation), { actual: operation || "unreported" }),
    check("pdf-trace:post-mutation-poppler-render", "trace", popplerAfterMerge, { actual: { mergeObserved: Boolean(merge), postMutationRenderObserved: popplerAfterMerge } }),
    check("pdf-trace:multi-source-audit-validation", "trace", auditAfterMerge, { actual: { mergeObserved: Boolean(merge), auditWithThreeInputsObserved: auditAfterMerge } }),
    gate("pdf-trace:no-ad-hoc-pdf-writer", "trace", !manualWriterPatterns.some((pattern) => pattern.test(commands.join("\n"))), { forbidden: manualWriterPatterns.map(String) }),
  ];
}

function ruledCrossPageTableTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const workflow = completedInvocation(commands, /officekit-ruled-cross-page-table-workflow\.mjs\b/i);
  const operation = auditOperation(audit);
  const genericOrManualPatterns = [
    /pdfplumber_extract\.py\b/i,
    /\bpdfplumber\.open\s*\(/i,
    /\bfind_tables\s*\(/i,
    /\bextract_tables?\s*\(/i,
    /\bPdfReader\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /^pdfplumber$/i.test(String(auditProvider(audit))), { expected: "pdfplumber", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^read-only$/i.test(String(auditSaveStrategy(audit))), { expected: "read-only", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:published-ruled-table-workflow", "trace", Boolean(workflow), { expected: "officekit-ruled-cross-page-table-workflow.mjs", actual: workflow ? commands[workflow.commandIndex] : "not observed" }),
    check("pdf-trace:workflow-preflight", "trace", audit?.preflight?.probeCompleted === true && audit?.preflight?.planCompleted === true && audit?.preflight?.sourceInspectionCompleted === true, { actual: audit?.preflight || "unreported" }),
    check("pdf-trace:typed-ruled-table-operation", "trace", /extract[-_ ]?ruled[-_ ]?cross[-_ ]?page[-_ ]?table/i.test(operation), { actual: operation || "unreported" }),
    check("pdf-trace:poppler-overlay-review", "trace", audit?.validation?.poppler?.status === "passed" && Array.isArray(audit?.validation?.poppler?.overlays), { actual: audit?.validation?.poppler || "unreported" }),
    gate("pdf-trace:no-generic-or-manual-table-route", "trace", !genericOrManualPatterns.some((pattern) => pattern.test(commandText)), { forbidden: genericOrManualPatterns.map(String) }),
  ];
}

function canonicalTableHeader(cells) {
  return Array.isArray(cells)
    ? cells.map((cell) => ({
      row: Number(cell?.row),
      column: Number(cell?.column),
      rowspan: Number(cell?.rowspan),
      colspan: Number(cell?.colspan),
      text: String(cell?.text || ""),
    }))
    : [];
}

function canonicalTableRows(rows) {
  return Array.isArray(rows)
    ? rows.map((row) => ({
      page: Number(row?.page),
      sourceRow: Number(row?.sourceRow),
      values: Array.isArray(row?.cells) ? row.cells.map((cell) => String(cell?.text || "")) : [],
    }))
    : [];
}

function finiteBbox(bbox, width, height) {
  return Array.isArray(bbox)
    && bbox.length === 4
    && bbox.every((value) => Number.isFinite(Number(value)))
    && Number(bbox[0]) >= 0
    && Number(bbox[1]) >= 0
    && Number(bbox[2]) > Number(bbox[0])
    && Number(bbox[3]) > Number(bbox[1])
    && Number(bbox[2]) <= width
    && Number(bbox[3]) <= height;
}

export function gradeRuledCrossPageTableEvidence({ evidence, audit, commands, item }) {
  const expected = item.grade?.machine || {};
  const result = evidence.json?.value || {};
  const table = result.table || {};
  const source = evidence.source || {};
  const expectedHeader = expected.header || [];
  const expectedRows = expected.rows || [];
  const expectedLabels = expected.columnLabels || [];
  const expectedFootnote = expected.footnote || "";
  const sourceTerms = expected.sourceTerms || [];
  const actualHeader = canonicalTableHeader(table.header);
  const actualRows = canonicalTableRows(table.dataRows);
  const expectedCanonicalRows = expectedRows.map((row) => ({ page: Number(row.page), sourceRow: Number(row.sourceRow), values: row.values.map(String) }));
  const expectedCsv = [
    ["page", "sourceRow", ...expectedLabels],
    ...expectedCanonicalRows.map((row) => [String(row.page), String(row.sourceRow), ...row.values]),
  ];
  const actualCsv = Array.isArray(evidence.csv?.rows) ? evidence.csv.rows : [];
  const pageFacts = new Map((source.pages || []).map((page) => [Number(page.page), page]));
  const segments = Array.isArray(table.segments) ? table.segments : [];
  const cellsBySegment = segments.flatMap((segment) => [
    ...(Array.isArray(segment?.headerCells) ? segment.headerCells : []),
    ...(Array.isArray(segment?.dataRows) ? segment.dataRows.flatMap((row) => Array.isArray(row?.cells) ? row.cells : []) : []),
  ].map((cell) => ({ cell, page: Number(segment?.page) })));
  const allCells = cellsBySegment.map(({ cell }) => cell);
  const geometryValid = allCells.length === expectedHeader.length * (expected.pageRange || []).length + expectedRows.length * Number(expected.expectedColumns || 0)
    && cellsBySegment.every(({ cell, page }) => {
      const sourcePage = pageFacts.get(page);
      return sourcePage
        && Number(cell?.page) === page
        && finiteBbox(cell?.bbox, Number(sourcePage.width), Number(sourcePage.height))
        && Number(cell?.confidence) >= 0.99
        && Array.isArray(cell?.confidenceReasons)
        && cell.confidenceReasons.includes("ruled-grid");
    });
  const repeatedHeaders = segments.length === (expected.pageRange || []).length
    && segments.every((segment, index) => Number(segment?.page) === Number(expected.pageRange?.[index])
      && JSON.stringify(canonicalTableHeader(segment?.headerCells)) === JSON.stringify(expectedHeader)
      && JSON.stringify((segment?.columns || []).map((column) => String(column?.label || ""))) === JSON.stringify(expectedLabels));
  const narrativeLeakage = /(?:NARRATIVE-(?:LEFT|RIGHT)|ROTATED-LABEL)/i.test(JSON.stringify({ table, csv: actualCsv }));
  const footnotes = Array.isArray(table.footnotes) ? table.footnotes : [];
  const overlays = audit?.validation?.poppler?.overlays;
  const visualPages = evidence.visual?.pages || [];
  const auditOutputCsv = audit?.outputs?.csv || {};
  const auditOutputJson = audit?.outputs?.json || audit?.output || {};
  const sourceIdentity = audit?.validation?.sourceIdentity || {};
  const auditChecks = audit?.validation?.ruledTable?.checks || [];
  return [
    check("pdf-machine:locked-source-profile", "machine", source.pageCount === expected.pageCount && sourceTerms.every((term) => Number(source.termCounts?.[term] || 0) > 0), { expected: { pageCount: expected.pageCount, sourceTerms }, actual: { pageCount: source.pageCount, termCounts: source.termCounts } }),
    check("pdf-machine:typed-table-schema-and-profile", "machine", result.schema === "office-kit.pdf-ruled-table.v1" && result.operation?.profile === "ruled-cross-page-v1" && result.operation?.type === "extract-ruled-cross-page-table" && result.operation?.tableTitle === expected.title && Number(result.operation?.expectedColumns) === Number(expected.expectedColumns) && Number(result.operation?.headerRows) === Number(expected.headerRows), { actual: { schema: result.schema, operation: result.operation } }),
    check("pdf-machine:page-range-repeated-header-and-columns", "machine", JSON.stringify(table.pageRange) === JSON.stringify(expected.pageRange) && repeatedHeaders && JSON.stringify(table.columnLabels) === JSON.stringify(expectedLabels) && JSON.stringify(actualHeader) === JSON.stringify(expectedHeader), { expected: { pageRange: expected.pageRange, header: expectedHeader, columnLabels: expectedLabels }, actual: { pageRange: table.pageRange, header: actualHeader, columnLabels: table.columnLabels } }),
    check("pdf-machine:exact-table-rows-and-csv", "machine", JSON.stringify(actualRows) === JSON.stringify(expectedCanonicalRows) && JSON.stringify(actualCsv) === JSON.stringify(expectedCsv), { expected: { rows: expectedCanonicalRows, csv: expectedCsv }, actual: { rows: actualRows, csv: actualCsv } }),
    check("pdf-machine:merged-spans-and-cell-geometry", "machine", geometryValid && actualHeader.length === expectedHeader.length && actualHeader.every((cell) => cell.rowspan >= 1 && cell.colspan >= 1), { actual: { cellCount: allCells.length, expectedCellCount: expectedHeader.length * (expected.pageRange || []).length + expectedRows.length * Number(expected.expectedColumns || 0), geometryValid } }),
    check("pdf-machine:footnote-and-no-narrative-leakage", "machine", footnotes.length === 1 && footnotes[0]?.text === expectedFootnote && Number(footnotes[0]?.page) === Number(expected.pageRange?.at(-1)) && narrativeLeakage === false, { expected: { footnote: expectedFootnote, page: expected.pageRange?.at(-1), noNarrativeLeakage: true }, actual: { footnotes, narrativeLeakage } }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), { actual: audit?.status || "unreported" }),
    check("pdf-visual:all-source-pages-rendered", "visual", evidence.visual?.renderer === "poppler-pdftoppm" && evidence.visual?.pageCount === expected.pageCount && visualPages.length === expected.pageCount && visualPages.every((page) => page.nonBlank && page.bytes > 1_000 && page.inkBBox), { actual: evidence.visual }),
    check("pdf-visual:overlay-review-bound-to-every-segment", "visual", audit?.validation?.poppler?.status === "passed" && audit?.validation?.poppler?.pageCount === expected.pageCount && Array.isArray(overlays) && overlays.length === expected.pageRange.length && overlays.every((overlay, index) => Number(overlay?.page) === Number(expected.pageRange[index]) && Number(overlay?.bytes) > 1_000 && Array.isArray(overlay?.tableBBox) && overlay.tableBBox.length === 4), { actual: audit?.validation?.poppler || "unreported" }),
    gate("pdf-security:source-and-output-provenance", "security", result.source?.sha256 === source.sha256 && auditSourceHash(audit) === source.sha256 && auditOutputHash(audit) === evidence.json?.sha256 && auditOutputJson?.sha256 === evidence.json?.sha256 && auditOutputCsv?.sha256 === evidence.csv?.sha256, { expected: { source: source.sha256, json: evidence.json?.sha256, csv: evidence.csv?.sha256 }, actual: { resultSource: result.source?.sha256, auditSource: auditSourceHash(audit), auditJson: auditOutputJson?.sha256, auditCsv: auditOutputCsv?.sha256 } }),
    gate("pdf-security:read-only-source-and-validated-profile", "security", auditSaveStrategy(audit) === "read-only" && audit?.savePolicy?.sourceOverwrite === false && sourceIdentity?.sourcePreserved === true && sourceIdentity?.before?.sha256 === source.sha256 && sourceIdentity?.after?.sha256 === source.sha256 && audit?.validation?.ruledTable?.passed === true && Array.isArray(auditChecks) && auditChecks.length >= 6 && auditChecks.every((entry) => entry?.passed === true), { actual: { savePolicy: audit?.savePolicy, sourceIdentity, ruledTable: audit?.validation?.ruledTable } }),
    gate("pdf-security:complete-typed-cell-proofs", "security", allCells.length > 0 && cellsBySegment.every(({ cell, page }) => Number.isInteger(cell?.page) && cell.page === page && Number.isInteger(cell?.row) && Number.isInteger(cell?.column) && Number.isInteger(cell?.rowspan) && Number.isInteger(cell?.colspan) && cell.rowspan >= 1 && cell.colspan >= 1 && finiteBbox(cell?.bbox, 612, 792)), { actual: { cells: allCells.length, cellsWithPage: allCells.filter((cell) => Number.isInteger(cell?.page)).length } }),
    ...ruledCrossPageTableTraceChecks(audit, commands),
  ];
}

function sortedRecords(records = []) {
  return [...records].sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)));
}

function sortedEntries(record = {}) {
  return Object.entries(record).sort(([left], [right]) => left.localeCompare(right));
}

function canonicalPath(value) {
  const resolved = path.resolve(String(value || ""));
  try {
    return realpathSync(resolved);
  } catch {
    return resolved;
  }
}

export function gradeMergeStampEvidence({ evidence, audit, commands, item }) {
  const output = evidence.output || {};
  const pageMap = evidence.pageMap || [];
  const visualPages = evidence.visual?.pages || [];
  const watermarkPages = new Set(item.grade.machine.watermarkPages || []);
  const expectedSequence = item.grade.machine.sequence || [];
  const actualSequence = pageMap.map((entry) => `${entry.source}:${entry.sourcePage}`);
  const geometryPreserved = pageMap.length === expectedSequence.length
    && pageMap.every((entry) => JSON.stringify(entry.sourceGeometry) === JSON.stringify(entry.outputGeometry));
  const watermarkPlacement = pageMap.every((entry) => (
    watermarkPages.has(entry.outputPage)
      ? entry.watermarkCount === 1 && entry.opacities?.some((value) => Math.abs(value - item.grade.machine.watermarkOpacity) <= 0.001)
      : entry.watermarkCount === 0
  ));
  const expectedNavigation = evidence.navigation?.expected || {};
  const actualNavigation = evidence.navigation?.actual || {};
  const navigationMatches = JSON.stringify(sortedRecords(expectedNavigation.outlines)) === JSON.stringify(sortedRecords(actualNavigation.outlines))
    && JSON.stringify(sortedEntries(expectedNavigation.namedDestinations)) === JSON.stringify(sortedEntries(actualNavigation.namedDestinations))
    && JSON.stringify(sortedRecords(expectedNavigation.internalLinks)) === JSON.stringify(sortedRecords(actualNavigation.internalLinks));
  const expectedInputs = Object.values(evidence.sources || {})
    .map((source) => ({ path: canonicalPath(source.path), bytes: source.bytes, sha256: source.sha256 }))
    .sort((left, right) => left.path.localeCompare(right.path));
  const actualInputs = (audit?.inputs || [])
    .map((source) => ({ path: canonicalPath(source.path), bytes: source.bytes, sha256: source.sha256 }))
    .sort((left, right) => left.path.localeCompare(right.path));
  const manifest = evidence.manifest?.value || {};
  const manifestSequence = (manifest.sequence || []).flatMap((segment) => {
    const sourcePageCount = evidence.sources?.[segment.source]?.pageCount || 0;
    const pages = segment.pages === "all" ? Array.from({ length: sourcePageCount }, (_, index) => index + 1) : segment.pages || [];
    return pages.map((page) => `${segment.source}:${page}`);
  });
  const manifestWatermark = manifest.watermarks?.[0] || {};
  const sourceTermsAbsent = Object.values(evidence.sources || {}).every((source) => source.termCounts?.[item.grade.machine.watermarkText] === 0);
  return [
    check("pdf-machine:page-count", "machine", output.pageCount === item.grade.machine.pageCount, { expected: item.grade.machine.pageCount, actual: output.pageCount }),
    check("pdf-machine:source-page-order", "machine", JSON.stringify(actualSequence) === JSON.stringify(expectedSequence), { expected: expectedSequence, actual: actualSequence }),
    check("pdf-machine:page-geometry-preserved", "machine", geometryPreserved, { actual: pageMap.map(({ outputPage, source, sourcePage, sourceGeometry, outputGeometry }) => ({ outputPage, source, sourcePage, sourceGeometry, outputGeometry })) }),
    check("pdf-machine:selective-watermark-and-opacity", "machine", watermarkPlacement, { expected: { pages: [...watermarkPages], text: item.grade.machine.watermarkText, opacity: item.grade.machine.watermarkOpacity }, actual: pageMap.map(({ outputPage, watermarkCount, opacities }) => ({ outputPage, watermarkCount, opacities })) }),
    check("pdf-machine:outline-count", "machine", actualNavigation.outlines?.length === item.grade.machine.outlines, { expected: item.grade.machine.outlines, actual: actualNavigation.outlines?.length }),
    check("pdf-machine:named-destination-count", "machine", Object.keys(actualNavigation.namedDestinations || {}).length === item.grade.machine.namedDestinations, { expected: item.grade.machine.namedDestinations, actual: Object.keys(actualNavigation.namedDestinations || {}).length }),
    check("pdf-machine:internal-link-count", "machine", actualNavigation.internalLinks?.length === item.grade.machine.internalLinks, { expected: item.grade.machine.internalLinks, actual: actualNavigation.internalLinks?.length }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), { actual: audit?.status || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual?.pageCount === item.grade.machine.pageCount && visualPages.length === item.grade.machine.pageCount && visualPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual?.renderer, pages: visualPages }),
    check("pdf-visual:non-watermarked-pages-stable", "visual", visualPages.filter((page) => !page.watermarkExpected).every((page) => page.pixelStable), { actual: visualPages.filter((page) => !page.watermarkExpected) }),
    check("pdf-visual:watermarked-pages-changed", "visual", visualPages.filter((page) => page.watermarkExpected).length === watermarkPages.size && visualPages.filter((page) => page.watermarkExpected).every((page) => !page.pixelStable && page.changedPixelsBBox), { actual: visualPages.filter((page) => page.watermarkExpected) }),
    gate("pdf-security:manifest-contract", "security", manifest.schema === "office-kit.pdf-merge-stamp.v1" && JSON.stringify(manifestSequence) === JSON.stringify(expectedSequence) && manifestWatermark.source === "report" && manifestWatermark.text === item.grade.machine.watermarkText && Math.abs(Number(manifestWatermark.opacity) - item.grade.machine.watermarkOpacity) <= 0.001, { actual: manifest }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === evidence.manifest?.sha256 && auditOutputHash(audit) === output.sha256, { expected: { source: evidence.manifest?.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    gate("pdf-security:all-source-bytes-bound", "security", JSON.stringify(actualInputs) === JSON.stringify(expectedInputs), { expected: expectedInputs, actual: actualInputs }),
    gate("pdf-security:navigation-resolved", "security", navigationMatches, { expected: expectedNavigation, actual: actualNavigation }),
    gate("pdf-security:single-revision-and-decodable", "security", output.startxrefCount === 1 && output.eofCount === 1 && output.decodedStreamErrors?.length === 0, { actual: { startxref: output.startxrefCount, eof: output.eofCount, decodedStreamErrors: output.decodedStreamErrors || [] } }),
    check("pdf-security:watermark-absent-from-sources", "security", sourceTermsAbsent, { actual: Object.fromEntries(Object.entries(evidence.sources || {}).map(([id, source]) => [id, source.termCounts?.[item.grade.machine.watermarkText]])) }),
    ...mergeStampTraceChecks(audit, commands),
  ];
}

export function gradeBoundedReplaceEvidence({ evidence, audit, commands, item }) {
  const expectedPageCount = item.grade.machine.pageCount;
  const oldTerms = item.grade.machine.notContains;
  const [newTerm, expectedNewCount] = Object.entries(item.grade.machine.contains)[0];
  const oldTerm = oldTerms[0];
  const targetPageNumber = item.grade.visual.allowedDiff[0].page;
  const output = evidence.output;
  const source = evidence.source;
  const renderPages = evidence.visual.pages || [];
  const nonTargetPages = renderPages.filter((page) => page.page !== targetPageNumber);
  const targetPage = renderPages.find((page) => page.page === targetPageNumber);
  const residue = oldTerms.every((term) => (
    output.termCounts?.[term] === 0
    && output.rawTermCounts?.[term] === 0
    && output.decodedStreamTermCounts?.[term] === 0
    && output.metadataTermCounts?.[term] === 0
  ));
  const auditStatus = String(audit?.status || "");
  return [
    check("pdf-machine:source-target-unique", "machine", source.termCounts?.[oldTerm] === 1 && source.pages?.[targetPageNumber - 1]?.termCounts?.[oldTerm] === 1, { expected: { term: oldTerm, count: 1, page: targetPageNumber }, actual: source.termCounts?.[oldTerm] }),
    check("pdf-machine:page-count", "machine", output.pageCount === expectedPageCount, { expected: expectedPageCount, actual: output.pageCount }),
    check("pdf-machine:replacement-count", "machine", output.termCounts?.[newTerm] === expectedNewCount && output.pages?.[targetPageNumber - 1]?.termCounts?.[newTerm] === expectedNewCount, { expected: { term: newTerm, count: expectedNewCount, page: targetPageNumber }, actual: output.termCounts?.[newTerm] }),
    check("pdf-machine:old-text-absent", "machine", output.termCounts?.[oldTerm] === 0, { expected: 0, actual: output.termCounts?.[oldTerm] }),
    check("pdf-machine:page-boxes-unchanged", "machine", samePageBoxes(source.pages, output.pages)),
    check("pdf-machine:font-geometry-unchanged", "machine", evidence.sourceStyle?.found && evidence.outputStyle?.found && JSON.stringify(evidence.sourceStyle.fonts) === JSON.stringify(evidence.outputStyle.fonts) && JSON.stringify(evidence.sourceStyle.sizes) === JSON.stringify(evidence.outputStyle.sizes) && sameBoundingBox(evidence.sourceStyle.bbox, evidence.outputStyle.bbox), { expected: { fonts: evidence.sourceStyle?.fonts, sizes: evidence.sourceStyle?.sizes, bbox: evidence.sourceStyle?.bbox }, actual: { fonts: evidence.outputStyle?.fonts, sizes: evidence.outputStyle?.sizes, bbox: evidence.outputStyle?.bbox } }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual.sourcePageCount === expectedPageCount && evidence.visual.outputPageCount === expectedPageCount && renderPages.length === expectedPageCount && renderPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual.renderer, pages: renderPages.length }),
    check("pdf-visual:non-target-pages-identical", "visual", nonTargetPages.length === expectedPageCount - 1 && nonTargetPages.every((page) => page.changedPixelsBBox === null), { actual: nonTargetPages.map(({ page, changedPixelsBBox }) => ({ page, changedPixelsBBox })) }),
    check("pdf-visual:target-diff-contained", "visual", Boolean(targetPage?.changedPixelsBBox) && targetPage?.changedWithinAllowedMask === true, { allowedMask: evidence.visual.allowedMask, actual: targetPage?.changedPixelsBBox || null }),
    gate("pdf-security:no-old-term-residue", "security", residue, { term: oldTerm, evidence: { text: output.termCounts?.[oldTerm], raw: output.rawTermCounts?.[oldTerm], decodedStreams: output.decodedStreamTermCounts?.[oldTerm], metadata: output.metadataTermCounts?.[oldTerm] } }),
    gate("pdf-security:all-streams-decodable", "security", source.decodedStreamErrors?.length === 0 && output.decodedStreamErrors?.length === 0, { actual: { source: source.decodedStreamErrors || [], output: output.decodedStreamErrors || [] } }),
    gate("pdf-security:single-revision", "security", output.startxrefCount === 1 && output.eofCount === 1, { expected: { startxref: 1, eof: 1 }, actual: { startxref: output.startxrefCount, eof: output.eofCount } }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === source.sha256 && auditOutputHash(audit) === output.sha256, { expected: { source: source.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    ...boundedTraceChecks(audit, commands),
  ];
}

function damagedXrefControl(audit) {
  const candidates = [
    audit?.controls?.unrecoverable,
    audit?.control?.unrecoverable,
    audit?.unrecoverable,
    audit?.validation?.unrecoverable,
    audit?.validation?.unrecoverableControl,
  ];
  return candidates.find((value) => value && typeof value === "object") || {};
}

function damagedXrefControlSourceHash(control, audit) {
  return control?.source?.sha256
    || control?.sourceSha256
    || control?.source_sha256
    || audit?.inputs?.find((input) => /unrecoverable/i.test(String(input?.role || input?.path || "")))?.sha256
    || "";
}

function normalizeQpdfInspection(record) {
  if (!record || typeof record !== "object") return record;
  return {
    ...record,
    status: record.status || record.checkStatus || record.qpdfStatus,
    exitCode: record.exitCode ?? record.checkExitCode ?? record.qpdfExitCode,
  };
}

function damagedXrefRepair(audit) {
  const qpdf = audit?.validation?.qpdf;
  if (qpdf && typeof qpdf === "object") {
    const namedFreshInspect = qpdf.freshInspect || qpdf.freshInspection;
    const stagedFreshInspect = qpdf.freshInspectCompleted === true
      ? (qpdf.after && typeof qpdf.after === "object"
        ? { ...qpdf.after, freshInspect: true }
        : (typeof qpdf.afterStatus === "string"
          ? {
            status: qpdf.afterStatus,
            exitCode: qpdf.afterExitCode,
            message: qpdf.afterMessage,
            freshInspect: true,
          }
          : undefined))
      : undefined;
    return {
      checkBefore: qpdf.before,
      qpdfWrite: qpdf.rewrite || audit?.warnings?.repairWrite,
      // A rewrite's own success claim is not evidence that the resulting
      // package is clean. Require a named fresh record, or an explicit
      // `freshInspectCompleted: true` staging marker plus its after record.
      checkAfter: normalizeQpdfInspection(namedFreshInspect || stagedFreshInspect),
      requiresFreshInspection: qpdf.freshInspectCompleted === true,
    };
  }
  const qpdfFreshInspect = audit?.validation?.qpdfFreshInspect;
  if (qpdfFreshInspect && typeof qpdfFreshInspect === "object") {
    return {
      checkBefore: audit?.warnings?.repairBefore,
      qpdfWrite: audit?.warnings?.repairWrite,
      checkAfter: normalizeQpdfInspection(qpdfFreshInspect),
      requiresFreshInspection: true,
    };
  }
  const freshInspection = audit?.validation?.freshInspection;
  if (freshInspection && typeof freshInspection === "object") {
    const repairEvidence = audit?.validation?.repairEvidence;
    return {
      // `freshInspection` is the required compact final-inspection record.
      // It must outrank a legacy `repairAfter` summary when both are present:
      // the latter may describe the same qpdf invocation rather than a fresh
      // post-write check.
      // Some agents keep the corresponding pre-repair warning evidence beside
      // it under `repairEvidence.before` rather than duplicating it at the
      // audit root. Preserve that evidence without treating the write itself
      // as a fresh inspection.
      checkBefore: repairEvidence?.checkBefore || repairEvidence?.before || audit?.warnings?.repairBefore,
      qpdfWrite: repairEvidence?.qpdfWrite || repairEvidence?.rewrite || audit?.warnings?.rewrite,
      checkAfter: normalizeQpdfInspection(freshInspection),
      requiresFreshInspection: true,
    };
  }
  const repairAfter = audit?.validation?.repairAfter;
  if (repairAfter && typeof repairAfter === "object") {
    return {
      checkBefore: audit?.validation?.repairBefore,
      qpdfWrite: audit?.warnings,
      checkAfter: normalizeQpdfInspection({
        status: repairAfter.qpdfStatus,
        freshInspect: repairAfter.freshInspect,
        exitCode: repairAfter.qpdfExitCode,
      }),
      requiresFreshInspection: true,
    };
  }
  const freshInspect = audit?.validation?.freshInspect;
  if (freshInspect && typeof freshInspect === "object") {
    return {
      checkBefore: audit?.warnings?.repairBefore,
      qpdfWrite: audit?.warnings?.repairWrite,
      // Agents commonly name this compact post-write record `freshInspect`.
      // Its completed flag remains mandatory: a clean-looking field attached
      // to the rewrite result itself is not independent evidence.
      checkAfter: freshInspect,
      requiresFreshInspection: true,
    };
  }
  const candidates = [
    audit?.repair,
    audit?.validation?.repair,
    audit?.validation?.repairEvidence,
    audit?.validation?.qpdfRepair,
  ];
  return candidates.find((value) => value && typeof value === "object") || audit || {};
}

function auditTextLines(value) {
  if (typeof value === "string") return [value];
  if (Array.isArray(value)) return value.flatMap(auditTextLines);
  if (!value || typeof value !== "object") return [];
  const textFields = new Set(["line", "lines", "jsonLines", "warnings", "message", "reason"]);
  return [
    ...auditTextLines(value.line),
    ...auditTextLines(value.lines),
    ...auditTextLines(value.jsonLines),
    ...auditTextLines(value.warnings),
    ...auditTextLines(value.message),
    ...auditTextLines(value.reason),
    // Warning records may be grouped by a semantic stage. Recurse through
    // nested records, but only collect approved textual evidence fields.
    ...Object.entries(value)
      .filter(([key, child]) => !textFields.has(key) && child && typeof child === "object")
      .flatMap(([, child]) => auditTextLines(child)),
  ];
}

function stableRecords(records = []) {
  return [...records].sort((left, right) => JSON.stringify(left).localeCompare(JSON.stringify(right)));
}

function qpdfRecoveryTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const inspect = completedQpdfProviderInvocation(commands, "inspect");
  const rewrite = completedQpdfProviderInvocation(commands, "rewrite");
  const afterRewrite = commandTextAfter(commands, rewrite);
  const outputInspect = hasQpdfProviderInvocation(afterRewrite, "inspect", rewrite?.bindings || []);
  const outputRender = /(?:\bpdftoppm\b|mupdf\.mjs["']?\s+render\b)/i.test(afterRewrite);
  const controlInspected = commands.some((command) => hasQpdfProviderInvocation(command, "inspect") && /unrecoverable\.pdf/i.test(String(command)))
    || /\bqpdf\s+--check\b[\s\S]*unrecoverable\.pdf/i.test(commandText);
  const manualWriterPatterns = [
    /\bPdfWriter\s*\(/,
    /\bcanvas\.Canvas\s*\(/,
    /\bwriter\.write\s*\(/,
    /\b(?:fitz|pymupdf)\.open\s*\(/i,
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /^qpdf$/i.test(String(auditProvider(audit))), { expected: "qpdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^rewrite$/i.test(String(auditSaveStrategy(audit))), { expected: "rewrite", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:inspect-before-repair", "trace", invocationBefore(inspect, rewrite), { actual: { inspect: inspect?.commandIndex ?? -1, rewrite: rewrite?.commandIndex ?? -1 } }),
    check("pdf-trace:typed-qpdf-repair", "trace", Boolean(rewrite) && /qpdf[-_ ]?rewrite|repair/i.test(operation), { actual: operation || "unreported" }),
    check("pdf-trace:unrecoverable-control-inspected", "trace", controlInspected, { expected: "explicit qpdf inspection of inputs/unrecoverable.pdf" }),
    check("pdf-trace:reinspect-and-render-output", "trace", outputInspect && outputRender, { actual: { outputInspect, outputRender } }),
    gate("pdf-trace:no-ad-hoc-pdf-rebuild", "trace", !manualWriterPatterns.some((pattern) => pattern.test(commandText)), { forbidden: manualWriterPatterns.map(String) }),
  ];
}

export function gradeDamagedXrefSafeRefusalEvidence({ evidence, audit, commands, finalMessage, outputEntries }) {
  const source = evidence.source || {};
  const control = damagedXrefControl(audit);
  const commandText = commands.join("\n");
  const noOutput = auditRefusalNoArtifactClaim(audit)
    && Array.isArray(outputEntries)
    && outputEntries.length === 1
    && outputEntries[0] === "audit.json";
  const diagnostic = `${finalMessage || ""}\n${audit?.reason || ""}\n${JSON.stringify(audit?.warnings || [])}`;
  const noRepair = !commands.some((command) => hasQpdfProviderInvocation(command, "rewrite"));
  return [
    check("pdf-machine:damaged-source-and-control-fixture", "machine", source.pageCount === 2 && evidence.sourceRawStartxrefIsZero === true && (evidence.sourceAttachments || []).length === 1 && evidence.unrecoverable?.hasControlCanary === true, { actual: { source, unrecoverable: evidence.unrecoverable } }),
    check("pdf-machine:safe-refusal-status", "machine", audit?.status === "failed_closed" && audit?.delivered_modified_pdf !== true, { actual: audit?.status || "unreported" }),
    check("pdf-machine:diagnostic", "machine", /qpdf|xref|repair|provider|unrecoverable/i.test(diagnostic), { actual: diagnostic || "unreported" }),
    check("pdf-visual:no-output-on-safe-refusal", "visual", noOutput, { actual: outputEntries || [] }),
    gate("pdf-security:both-inputs-bound", "security", auditSourceHash(audit) === source.sha256 && damagedXrefControlSourceHash(control, audit) === evidence.unrecoverable?.sha256, { expected: { recoverable: source.sha256, unrecoverable: evidence.unrecoverable?.sha256 }, actual: { recoverable: auditSourceHash(audit) || "unreported", unrecoverable: damagedXrefControlSourceHash(control, audit) || "unreported" } }),
    gate("pdf-security:no-partial-output", "security", noOutput, { actual: { audit: audit?.output, outputEntries } }),
    check("pdf-trace:provider-declared", "trace", Boolean(String(auditProvider(audit)).trim()), { actual: auditProvider(audit) || "unreported" }),
    check("pdf-trace:source-inspection", "trace", commands.some((command) => hasQpdfProviderInvocation(command, "probe") || hasQpdfProviderInvocation(command, "inspect")) || /\bqpdf\b/i.test(commandText), { actual: commandText || "unreported" }),
    check("pdf-trace:no-repair-after-refusal", "trace", noRepair, { actual: commandText || "unreported" }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
  ];
}

export function gradeDamagedXrefRecoveryEvidence({ evidence, audit, commands, item }) {
  const source = evidence.source || {};
  const output = evidence.output || {};
  const visual = evidence.visual || {};
  const pages = visual.pages || [];
  const repair = damagedXrefRepair(audit);
  const control = damagedXrefControl(audit);
  const expectedAttachments = stableRecords(evidence.sourceAttachments || []);
  const actualAttachments = stableRecords(evidence.outputAttachments || []);
  const textCanaries = ["RECOVERABLE-XREF-PAGE-ONE", "RECOVERABLE-XREF-PAGE-TWO"];
  const sourceFixtureComplete = source.pageCount === 2
    && evidence.sourceRawStartxrefIsZero === true
    && textCanaries.every((term, index) => source.termCounts?.[term] === 1 && source.pages?.[index]?.termCounts?.[term] === 1)
    && expectedAttachments.length === 1
    && expectedAttachments[0]?.displayName === "repair-evidence.txt"
    && expectedAttachments[0]?.bytes === 97
    && expectedAttachments[0]?.sha256 === "f3847151df00468cce51c49a00cec7b5c35c24421160e310cbde0e90148b5550"
    && evidence.unrecoverable?.hasPdfHeader === true
    && evidence.unrecoverable?.hasControlCanary === true
    && evidence.unrecoverable?.hasXrefOrTrailer === false;
  const outputCanariesStable = textCanaries.every((term, index) => output.termCounts?.[term] === 1 && output.pages?.[index]?.termCounts?.[term] === 1)
    && output.decodedStreamTermCounts?.["OFFICEKIT-XREF-ATTACHMENT-CANARY"] === 1;
  const warningLines = auditTextLines([
    repair?.checkBefore,
    repair?.qpdfWrite,
    repair?.warnings,
    audit?.warnings,
  ]).join("\n");
  const controlStatus = String(control?.status || control?.result || control?.decision || "").toLowerCase();
  const controlRejected = (/reject|fail|error|unrecoverable|refus/.test(controlStatus) || control?.qpdfAccepted === false || control?.qpdfRejected === true || control?.accepted === false || control?.rejected === true || control?.refused === true)
    && damagedXrefControlSourceHash(control, audit) === evidence.unrecoverable?.sha256
    && (control?.output === null || control?.outputPath === null || control?.artifactWritten === false || control?.published === false || control?.outputGenerated === false || control?.outputPresent === false || control?.pseudoRepairProduced === false || control?.pseudoRepairGenerated === false);
  const qpdfOutputClean = repair?.checkAfter?.status === "clean"
    && (!repair?.requiresFreshInspection || repair?.checkAfter?.freshInspect === true || repair?.checkAfter?.completed === true);
  return [
    check("pdf-machine:damaged-source-and-control-fixture", "machine", sourceFixtureComplete, { actual: { source: { pageCount: source.pageCount, startxrefIsZero: evidence.sourceRawStartxrefIsZero, attachments: expectedAttachments }, unrecoverable: evidence.unrecoverable } }),
    check("pdf-machine:qpdf-output-clean", "machine", qpdfOutputClean, { actual: repair?.checkAfter || "unreported" }),
    check("pdf-machine:page-count-and-geometry-preserved", "machine", output.pageCount === source.pageCount && samePageBoxes(source.pages, output.pages), { actual: { source: source.pageCount, output: output.pageCount } }),
    check("pdf-machine:document-attachment-preserved", "machine", JSON.stringify(actualAttachments) === JSON.stringify(expectedAttachments), { expected: expectedAttachments, actual: actualAttachments }),
    check("pdf-machine:unrecoverable-control-rejected", "machine", controlRejected, { expected: { sourceSha256: evidence.unrecoverable?.sha256, noOutput: true }, actual: control || "unreported" }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), { actual: audit?.status || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", visual.sourcePageCount === source.pageCount && visual.outputPageCount === source.pageCount && pages.length === source.pageCount && pages.every((page) => page.sameDimensions && page.nonBlank), { renderer: visual.renderer, pages }),
    check("pdf-visual:all-pages-pixel-stable", "visual", pages.length === source.pageCount && pages.every((page) => page.pixelStable === true), { actual: pages.map(({ page, changedPixelsBBox }) => ({ page, changedPixelsBBox })) }),
    gate("pdf-security:repair-warnings-retained", "security", /reconstruct cross-reference/i.test(warningLines) && /file is damaged/i.test(warningLines), { actual: warningLines || "unreported" }),
    gate("pdf-security:no-raster-reconstruction", "security", outputCanariesStable && JSON.stringify(actualAttachments) === JSON.stringify(expectedAttachments) && output.decodedStreamErrors?.length === 0, { actual: { text: output.termCounts, decoded: output.decodedStreamTermCounts, attachments: actualAttachments, decodedErrors: output.decodedStreamErrors || [] } }),
    gate("pdf-security:single-clean-revision-and-provenance", "security", output.startxrefCount === 1 && output.eofCount === 1 && auditSourceHash(audit) === source.sha256 && auditOutputHash(audit) === output.sha256 && output.sha256 !== source.sha256, { expected: { source: source.sha256, output: output.sha256 }, actual: { startxref: output.startxrefCount, eof: output.eofCount, source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    ...qpdfRecoveryTraceChecks(audit, commands),
  ];
}

export function gradeSourceBoundHighlightEvidence({ evidence, audit, commands, item }) {
  const source = evidence.source || {};
  const output = evidence.output || {};
  const targetText = item.grade.machine.text;
  const targetPageNumber = Number(item.grade.machine.targetPage);
  const expectedColor = item.grade.machine.color || [1, 1, 0];
  const sourcePage = source.pages?.[targetPageNumber - 1] || {};
  const renderPages = evidence.visual?.pages || [];
  const targetRender = renderPages.find((page) => page.page === targetPageNumber);
  const nonTargetRenders = renderPages.filter((page) => page.page !== targetPageNumber);
  const outputHighlights = evidence.outputHighlights || [];
  const sourceHighlights = evidence.sourceHighlights || [];
  const matchingHighlights = outputHighlights.filter((highlight) => (
    highlight.page === targetPageNumber
    && highlight.contents === item.grade.machine.contents
    && highlight.author === item.grade.machine.author
    && highlight.subject === item.grade.machine.subject
    && highlightColorMatches(highlight.color, expectedColor)
  ));
  const operation = auditOperationRecord(audit);
  const expectedPage = operation?.expectedPage;
  const expectedPageBBox = [0, 0, Number(sourcePage.width), Number(sourcePage.height)];
  // The provider's applied-operation record intentionally omits the redundant
  // source hash; its enclosing audit metadata binds that exact input bytes.
  // Require the applied source-bound type and page snapshot here, then verify
  // the audit source/output hashes in the provenance gate below.
  const sourceBoundSnapshot = operation?.type === "add_text_highlight"
    && sameBoundingBox(expectedPage?.bbox || [], expectedPageBBox)
    && Number(expectedPage?.rotation) === Number(sourcePage.rotation);
  const unsupportedCallerGeometry = ["bbox", "rect", "quadPoints", "quads", "point"].some((field) => Object.hasOwn(operation || {}, field));
  const sourceTextStable = output.termCounts?.[targetText] === 1
    && output.pages?.[targetPageNumber - 1]?.termCounts?.[targetText] === 1;
  const contentGeometryStable = evidence.sourceTarget?.found
    && evidence.outputTarget?.found
    && JSON.stringify(evidence.sourceTarget.fonts) === JSON.stringify(evidence.outputTarget.fonts)
    && JSON.stringify(evidence.sourceTarget.sizes) === JSON.stringify(evidence.outputTarget.sizes)
    && sameBoundingBox(evidence.sourceTarget.bbox, evidence.outputTarget.bbox);
  return [
    check("pdf-machine:source-target-unique", "machine", source.termCounts?.[targetText] === 1 && source.pages?.[targetPageNumber - 1]?.termCounts?.[targetText] === 1, { expected: { text: targetText, count: 1, page: targetPageNumber }, actual: source.termCounts?.[targetText] }),
    check("pdf-machine:page-count-and-geometry", "machine", output.pageCount === item.grade.machine.pageCount && samePageBoxes(source.pages, output.pages), { expected: item.grade.machine.pageCount, actual: output.pageCount }),
    check("pdf-machine:original-text-stable", "machine", sourceTextStable && contentGeometryStable, { actual: { termCount: output.termCounts?.[targetText], targetPageCount: output.pages?.[targetPageNumber - 1]?.termCounts?.[targetText], sourceTarget: evidence.sourceTarget, outputTarget: evidence.outputTarget } }),
    check("pdf-machine:one-native-highlight", "machine", sourceHighlights.length === 0 && outputHighlights.length === 1 && matchingHighlights.length === 1, { expected: { page: targetPageNumber, contents: item.grade.machine.contents, author: item.grade.machine.author, subject: item.grade.machine.subject, color: expectedColor }, actual: outputHighlights }),
    check("pdf-machine:highlight-quads-bound-to-target-text", "machine", matchingHighlights.length === 1 && highlightQuadMatchesSourceText(matchingHighlights[0], evidence.sourceTarget, Number(sourcePage.height)), { sourceText: evidence.sourceTarget, actual: matchingHighlights[0]?.quadPoints || [] }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), { actual: audit?.status || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual?.sourcePageCount === item.grade.machine.pageCount && evidence.visual?.outputPageCount === item.grade.machine.pageCount && renderPages.length === item.grade.machine.pageCount && renderPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual?.renderer, pages: renderPages }),
    check("pdf-visual:non-target-pages-identical", "visual", nonTargetRenders.length === item.grade.machine.pageCount - 1 && nonTargetRenders.every((page) => page.changedPixelsBBox === null), { actual: nonTargetRenders.map(({ page, changedPixelsBBox }) => ({ page, changedPixelsBBox })) }),
    check("pdf-visual:highlight-diff-contained", "visual", Boolean(targetRender?.changedPixelsBBox) && targetRender?.changedWithinAllowedMask === true, { allowedMask: evidence.visual?.allowedMask, actual: targetRender?.changedPixelsBBox || null }),
    gate("pdf-security:source-bound-operation", "security", sourceBoundSnapshot && !unsupportedCallerGeometry, { expected: { sourceSha256: source.sha256, expectedPage: { bbox: expectedPageBBox, rotation: sourcePage.rotation }, forbiddenCallerGeometry: true }, actual: operation || "unreported" }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === source.sha256 && auditOutputHash(audit) === output.sha256 && output.sha256 !== source.sha256, { expected: { source: source.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    gate("pdf-security:single-revision-and-decodable", "security", source.decodedStreamErrors?.length === 0 && output.decodedStreamErrors?.length === 0 && output.startxrefCount === 1 && output.eofCount === 1, { actual: { sourceErrors: source.decodedStreamErrors || [], outputErrors: output.decodedStreamErrors || [], startxref: output.startxrefCount, eof: output.eofCount } }),
    ...sourceBoundHighlightTraceChecks(audit, commands),
  ];
}

export function gradeOverflowRefusalEvidence({ evidence, audit, commands, finalMessage, item }) {
  const geometry = evidence.geometry || {};
  const auditGeometry = audit?.geometry_evidence || audit?.geometryEvidence || {};
  const operation = String(auditOperation(audit));
  const diagnostics = item.grade.machine.diagnosticTerms || [];
  const diagnosticText = `${finalMessage}\n${audit?.reason || ""}`.toLowerCase();
  const auditWidth = Number(auditGeometry.replacement_text_width_pt ?? auditGeometry.replacementWidth);
  const auditAvailable = Number(auditGeometry.available_width_from_text_origin_pt ?? auditGeometry.availableWidth);
  const auditFits = auditGeometry.fits_outer_cell_without_wrap ?? auditGeometry.fitsObservedTextArea ?? auditGeometry.fits;
  return [
    check("pdf-machine:source-target-unique", "machine", evidence.source.termCounts?.Approved === 1 && evidence.source.pages?.[0]?.termCounts?.Approved === 1, { expected: 1, actual: evidence.source.termCounts?.Approved }),
    check("pdf-machine:independent-overflow-proof", "machine", sameNumber(geometry.targetRectangle?.width, 70) && geometry.fits === false && geometry.replacementWidth > geometry.availableWidth, { actual: geometry }),
    check("pdf-machine:audit-overflow-proof", "machine", Number.isFinite(auditWidth) && Number.isFinite(auditAvailable) && auditWidth > auditAvailable && auditFits === false, { actual: { replacementWidth: auditWidth, availableWidth: auditAvailable, fits: auditFits } }),
    check("pdf-machine:safe-refusal-status", "machine", audit?.status === "failed_closed" && audit?.delivered_modified_pdf !== true, { actual: audit?.status || "unreported" }),
    check("pdf-machine:diagnostic", "machine", diagnostics.some((term) => diagnosticText.includes(String(term).toLowerCase())), { expectedAny: diagnostics }),
    gate("pdf-security:no-partial-output", "security", audit?.validation?.no_partial_modified_pdf_in_outputs === true || audit?.noPartialOutput === true, { actual: audit?.validation?.no_partial_modified_pdf_in_outputs ?? audit?.noPartialOutput ?? "unreported" }),
    gate("pdf-security:no-mutation-claimed", "security", /(?:no (?:pdf )?mutation|not performed|未(?:执行|修改)|没有修改)/i.test(operation), { actual: operation || "unreported" }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === evidence.source.sha256, { expected: evidence.source.sha256, actual: auditSourceHash(audit) || "unreported" }),
    ...overflowTraceChecks(audit, commands),
  ];
}

function boundaryFixtureComplete(evidence, item) {
  if (item.id === "pdf-encrypted-owner-policy-boundary") {
    return evidence.encrypted === true
      && evidence.userPasswordAccepted === true
      && evidence.encryption?.version >= 5
      && evidence.encryption?.revision >= 5
      && evidence.encryption?.copyAllowed === false
      && evidence.hasEmbeddedFiles === true
      && evidence.hasAcroForm === true
      && evidence.hasSignatureField === true;
  }
  if (item.id === "pdf-annotation-reply-resolve-boundary") {
    return evidence.root?.present === true
      && evidence.root?.author === "Reviewer A"
      && evidence.root?.creationDate === "D:20260728090000Z"
      && evidence.root?.state === "/Accepted"
      && evidence.root?.stateModel === "/Review"
      && evidence.replyCount === 1
      && evidence.replyStates?.[0]?.author === "Data steward"
      && evidence.replyStates?.[0]?.state === "/Accepted"
      && evidence.replyStates?.[0]?.stateModel === "/Review"
      && evidence.popupCount === 1
      && evidence.highlightCount === 1;
  }
  if (item.id === "pdf-auto-pdfua-overclaim-refusal") {
    return evidence.hasStructTreeRoot === false
      && evidence.imageCount >= 1
      && evidence.hasTwoColumnCanaries === true
      && evidence.pageCount >= 1;
  }
  if (item.id === "pdf-mixed-scan-ocr-boundary") {
    return evidence.pageCount === 8
      && evidence.language === "zh-CN"
      && evidence.metadata?.title === "Mixed scan OCR boundary fixture"
      && /OFFICEKIT-MIXED-SCAN-OCR-BOUNDARY/.test(evidence.metadata?.keywords || "")
      && JSON.stringify(evidence.imageOnlyPages) === JSON.stringify([1, 2, 3, 4, 5, 6])
      && JSON.stringify(evidence.imageCounts) === JSON.stringify([1, 1, 1, 1, 1, 1, 0, 0])
      && evidence.scanImageCanariesMatch === true
      && evidence.bornDigitalCanaries === true
      && evidence.requires?.pixelEncodedUpsideDown === true
      && evidence.requires?.pixelEncodedSideways === true
      && JSON.stringify(evidence.requires?.pixelEncodedDeskewDegrees) === JSON.stringify([3, -2])
      && evidence.requires?.mixedImageOnlyAndBornDigitalPages === true;
  }
  if (item.id === "pdf-dynamic-xfa-boundary") {
    return evidence.hasXfaPackets === true
      && evidence.needsRendering === true
      && evidence.hasRepeatSubform === true
      && evidence.hasFormCalc === true
      && evidence.hasJavaScript === true;
  }
  if (item.id === "pdf-print-production-boundary") {
    return evidence.hasOutputIntent === true
      && evidence.hasOcg === true
      && evidence.hasDeviceN === true
      && evidence.hasSeparation === true
      && evidence.overprint === true
      && evidence.transparency === true;
  }
  if (item.id === "pdf-richmedia-opaque-preservation") {
    const hashes = Object.values(evidence.opaqueStreamSha256 || {});
    return evidence.pageCount === 2
      && evidence.threeDAnnotationCount === 1
      && evidence.richMediaAnnotationCount === 1
      && evidence.hasThreeDStream === true
      && evidence.hasThreeDDefaultView === true
      && evidence.hasThreeDActivation === true
      && evidence.hasRichMediaContent === true
      && evidence.hasRichMediaSettings === true
      && evidence.hasRichMediaAssets === true
      && evidence.hasRichMediaConfiguration === true
      && evidence.hasRichMediaActivation === true
      && evidence.hasJavaScript === true
      && hashes.length === 2
      && hashes.every((hash) => typeof hash === "string" && /^[a-f0-9]{64}$/i.test(hash));
  }
  if (item.id === "pdf-docmdp-forbidden-title-edit") {
    return evidence.pageCount >= 1
      && evidence.metadataTitle === "Final"
      && evidence.hasPerms === true
      && evidence.hasDocMDP === true
      && evidence.byteRangeValid === true
      && evidence.cmsContentsPresent === true
      && evidence.docMDPTransformCount === 1
      && evidence.docMDPPermission === 1;
  }
  return false;
}

export function gradeBoundaryRefusalEvidence({ evidence, audit, commands, finalMessage, item, outputEntries }) {
  const diagnostics = item.grade?.machine?.diagnosticTerms || [];
  const diagnosticText = `${finalMessage}\n${audit?.reason || ""}\n${JSON.stringify(audit?.warnings || [])}`.toLowerCase();
  const operationRecord = auditOperationRecord(audit) || {};
  const actualOutputs = Array.isArray(outputEntries) ? outputEntries : null;
  return [
    check("pdf-machine:source-boundary-fixture", "machine", boundaryFixtureComplete(evidence, item), { actual: evidence }),
    check("pdf-machine:safe-refusal-status", "machine", audit?.status === "failed_closed" && audit?.delivered_modified_pdf !== true, { actual: audit?.status || "unreported" }),
    check("pdf-machine:diagnostic", "machine", diagnostics.some((term) => diagnosticText.includes(String(term).toLowerCase())), { expectedAny: diagnostics }),
    check("pdf-machine:refusal-operation", "machine", auditRefusalOperationUnexecuted(audit), { expected: "an explicit requested operation with a verified no-mutation result", actual: operationRecord || "unreported" }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === evidence.source?.sha256, { expected: evidence.source?.sha256, actual: auditSourceHash(audit) || "unreported" }),
    gate("pdf-security:no-partial-output", "security", auditRefusalNoArtifactClaim(audit) && (actualOutputs === null || actualOutputs.length === 1 && actualOutputs[0] === "audit.json"), { expected: "audit-only outputs plus a failed-closed null-output claim", actual: { auditOutput: audit?.output, outputEntries: actualOutputs } }),
    ...boundaryRefusalTraceChecks(audit, commands, item),
  ];
}

export function gradeAcroFormEvidence({ evidence, audit, commands, item }) {
  const expectedFields = item.grade.machine.fields;
  const sourceFields = evidence.sourceForm.fields || {};
  const outputFields = evidence.outputForm.fields || {};
  const sourceWidgets = evidence.sourceForm.widgets || [];
  const outputWidgets = evidence.outputForm.widgets || [];
  const visualPages = evidence.visual.pages || [];
  const widgetChanges = evidence.visual.widgetChanges || [];
  const expectedOutputValue = (name, value) => outputFields[name]?.fieldType === "/Btn" && value
    ? `/${String(value).replace(/^\//, "")}`
    : String(value);
  const sourceFixtureComplete = Object.keys(expectedFields).every((name) => (
    sourceFields[name]
    && sourceFields[name].value === ""
    && sourceFields[name].readOnly === false
  ))
    && sourceFields.company_type?.states?.includes("/LLC")
    && sourceFields.company_type?.states?.includes("/Corporation")
    && sourceFields.terms_ack?.fieldType === "/Btn"
    && sourceFields.terms_ack?.value === "/Yes"
    && sourceFields.terms_ack?.readOnly === false;
  const valuesExact = Object.entries(expectedFields).every(([name, value]) => (
    outputFields[name]?.value === expectedOutputValue(name, value)
  ));
  const topologyPreserved = sourceWidgets.length === outputWidgets.length && sourceWidgets.every((source, index) => {
    const output = outputWidgets[index];
    return output
      && source.page === output.page
      && source.name === output.name
      && source.fieldType === output.fieldType
      && sameBoundingBox(source.rect, output.rect)
      && output.readOnly === false;
  });
  const appearancesPresent = outputWidgets.length > 0 && outputWidgets.every((widget) => widget.appearancePresent === true);
  const companyWidgets = outputWidgets.filter((widget) => widget.name === "company_type");
  const radioStateCorrect = companyWidgets.length === 2
    && companyWidgets.filter((widget) => widget.selectedState === "/LLC").length === 1
    && companyWidgets.filter((widget) => widget.selectedState === "/Off").length === 1;
  const sourceCheckbox = sourceWidgets.find((widget) => widget.name === "terms_ack");
  const outputCheckbox = outputWidgets.find((widget) => widget.name === "terms_ack");
  const checkboxPreserved = outputFields.terms_ack?.value === sourceFields.terms_ack?.value
    && sourceCheckbox?.selectedState === "/Yes"
    && outputCheckbox?.selectedState === sourceCheckbox.selectedState;
  const changedWidgets = widgetChanges.filter((widget) => widget.expectedChange);
  const untouchedWidgets = widgetChanges.filter((widget) => !widget.expectedChange);
  const expectedChangedWidgetCount = Object.entries(expectedFields).filter(([, value]) => String(value).length > 0).length;
  const expectedUntouchedWidgetCount = sourceWidgets.length - expectedChangedWidgetCount;
  const sensitiveFieldsBlank = ["tin", "signature"].every((name) => (
    outputFields[name]?.value === ""
    && outputFields[name]?.defaultValue === ""
    && widgetChanges.filter((widget) => widget.name === name).every((widget) => widget.changedPixelsBBox === null)
  ));
  const auditStatus = String(audit?.status || "");
  return [
    check("pdf-machine:source-form-fixture-complete", "machine", sourceFixtureComplete, { actual: sourceFields }),
    check("pdf-machine:page-count-and-boxes-unchanged", "machine", evidence.output.pageCount === evidence.source.pageCount && samePageBoxes(evidence.source.pages, evidence.output.pages), { actual: { source: evidence.source.pageCount, output: evidence.output.pageCount } }),
    check("pdf-machine:field-values-exact", "machine", valuesExact, { expected: expectedFields, actual: Object.fromEntries(Object.entries(outputFields).map(([name, field]) => [name, field.value])) }),
    check("pdf-machine:widget-topology-preserved", "machine", topologyPreserved, { actual: { sourceWidgets: sourceWidgets.length, outputWidgets: outputWidgets.length } }),
    check("pdf-machine:appearances-present", "machine", appearancesPresent && evidence.outputForm.needAppearances === false, { actual: { appearancesPresent, needAppearances: evidence.outputForm.needAppearances } }),
    check("pdf-machine:radio-appearance-state", "machine", radioStateCorrect, { actual: companyWidgets.map(({ appearanceStates, selectedState }) => ({ appearanceStates, selectedState })) }),
    check("pdf-machine:checkbox-state-preserved", "machine", checkboxPreserved, { actual: { source: sourceCheckbox?.selectedState, output: outputCheckbox?.selectedState, fieldValue: outputFields.terms_ack?.value } }),
    check("pdf-machine:interactive-form-preserved", "machine", evidence.outputForm.acroFormPresent === true && evidence.outputForm.fieldTreeRoots === evidence.sourceForm.fieldTreeRoots && evidence.outputForm.fieldTreeRoots > 0, { actual: { acroFormPresent: evidence.outputForm.acroFormPresent, fieldTreeRoots: evidence.outputForm.fieldTreeRoots } }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual.sourcePageCount === evidence.source.pageCount && evidence.visual.outputPageCount === evidence.source.pageCount && visualPages.length === evidence.source.pageCount && visualPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual.renderer, pages: visualPages.length }),
    check("pdf-visual:changes-contained-to-target-widgets", "visual", visualPages.every((page) => page.changedOnlyWithinAllowedMasks), { actual: visualPages.map(({ page, changedOutsideAllowedMasksBBox }) => ({ page, changedOutsideAllowedMasksBBox })) }),
    check("pdf-visual:filled-values-visible", "visual", changedWidgets.length === expectedChangedWidgetCount && changedWidgets.every((widget) => widget.changedPixelsBBox !== null && widget.changedInteriorPixelsBBox !== null), { expected: expectedChangedWidgetCount, actual: changedWidgets }),
    check("pdf-visual:untouched-widgets-stable", "visual", untouchedWidgets.length === expectedUntouchedWidgetCount && untouchedWidgets.every((widget) => widget.changedPixelsBBox === null), { expected: expectedUntouchedWidgetCount, actual: untouchedWidgets }),
    gate("pdf-security:incremental-prefix-preserved", "security", evidence.originalPrefixPreserved === true, { actual: evidence.originalPrefixPreserved }),
    gate("pdf-security:single-appended-revision", "security", evidence.output.startxrefCount === evidence.source.startxrefCount + 1 && evidence.output.eofCount === evidence.source.eofCount + 1, { expected: { startxref: evidence.source.startxrefCount + 1, eof: evidence.source.eofCount + 1 }, actual: { startxref: evidence.output.startxrefCount, eof: evidence.output.eofCount } }),
    gate("pdf-security:sensitive-fields-blank", "security", sensitiveFieldsBlank, { actual: { tin: outputFields.tin, signature: outputFields.signature } }),
    gate("pdf-security:all-streams-decodable", "security", evidence.source.decodedStreamErrors?.length === 0 && evidence.output.decodedStreamErrors?.length === 0, { actual: { source: evidence.source.decodedStreamErrors || [], output: evidence.output.decodedStreamErrors || [] } }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === evidence.source.sha256 && auditOutputHash(audit) === evidence.output.sha256, { expected: { source: evidence.source.sha256, output: evidence.output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    ...acroformTraceChecks(audit, commands),
  ];
}

export function gradeCertifiedDocMdpP2FillEvidence({ evidence, audit, commands, item }) {
  const targetName = Object.keys(item.grade.machine.field || {})[0] || "ApprovedAmount";
  const targetValue = String(item.grade.machine.field?.[targetName] || "12500.00");
  const lockedName = "LockedAmount";
  const lockedValue = "LOCKED-9000";
  const sourceFields = evidence.sourceForm?.fields || {};
  const outputFields = evidence.outputForm?.fields || {};
  const sourceWidgets = evidence.sourceForm?.widgets || [];
  const outputWidgets = evidence.outputForm?.widgets || [];
  const sourceTarget = sourceFields[targetName];
  const outputTarget = outputFields[targetName];
  const sourceLocked = sourceFields[lockedName];
  const outputLocked = outputFields[lockedName];
  const sourceTargetWidgets = sourceWidgets.filter((widget) => widget.name === targetName);
  const outputTargetWidgets = outputWidgets.filter((widget) => widget.name === targetName);
  const visualPages = evidence.visual?.pages || [];
  const widgetChanges = evidence.visual?.widgetChanges || [];
  const targetWidgetChanges = widgetChanges.filter((widget) => widget.name === targetName);
  const sourceSignature = evidence.sourceSignature || {};
  const outputSignature = evidence.outputSignature || {};
  const sourceByteRange = sourceSignature.signature?.byteRange || {};
  const outputByteRange = outputSignature.signature?.byteRange || {};
  const sourceDocMdp = sourceSignature.docMDP || {};
  const outputDocMdp = outputSignature.docMDP || {};
  const sourceFieldMdp = sourceSignature.fieldMDP || {};
  const outputFieldMdp = outputSignature.fieldMDP || {};
  const auditPostflight = audit?.signature?.postflight || {};
  const auditPreflight = audit?.signature?.preflight || {};
  const auditField = audit?.field || {};
  const auditControlledFinalisation = audit?.schema === "office-kit.pyhanko-certified-form-fill.v1"
    && audit?.ok === true
    && audit?.operationCompleted === true
    && audit?.operation === "fill-certified-docmdp-p2-text-field"
    && audit?.silentFallback === false
    && audit?.trustRoot?.sha256
    && /^[a-f0-9]{64}$/i.test(String(audit.trustRoot.sha256))
    && audit?.savePolicy?.strategy === "incremental"
    && audit?.savePolicy?.sourcePrefixPreserved === true
    && Number(audit?.savePolicy?.revisionsAfter) === Number(audit?.savePolicy?.revisionsBefore) + 1
    && auditField?.target === targetName
    && auditField?.value === targetValue
    && auditField?.locked?.name === lockedName
    && auditField?.locked?.value === lockedValue
    && auditField?.nonTargetFieldsUnchanged === true
    && audit?.transaction?.distinctOutput === true
    && audit?.transaction?.noReplace === true
    && audit?.transaction?.outputPublishedAtomically === true
    && auditPreflight?.fieldName === "Certification"
    && auditPreflight?.coverage === "entire-file"
    && auditPreflight?.modificationLevel === "none"
    && auditPreflight?.docMDP?.permission === "fill-forms"
    && auditPreflight?.fieldMDP?.action === "include"
    && JSON.stringify(auditPreflight?.fieldMDP?.fields) === JSON.stringify([lockedName])
    && auditPostflight?.fieldName === "Certification"
    && auditPostflight?.coverage === "entire-revision"
    && auditPostflight?.modificationLevel === "form-filling"
    && auditPostflight?.cryptographicallyValid === true
    && auditPostflight?.intact === true
    && auditPostflight?.trusted === true
    && auditPostflight?.bottomLine === true
    && auditPostflight?.docMDPCompliant === true
    && auditPostflight?.docMDP?.permission === "fill-forms"
    && auditPostflight?.fieldMDP?.action === "include"
    && JSON.stringify(auditPostflight?.fieldMDP?.fields) === JSON.stringify([lockedName])
    && JSON.stringify(auditPostflight?.changedFormFields) === JSON.stringify([targetName]);
  const sourceFixtureComplete = sourceSignature.signature?.present === true
    && sourceByteRange.validSegments === true
    && sourceByteRange.coversEntireFile === true
    && sourceDocMdp.transformCount === 1
    && sourceDocMdp.permission === 2
    && sourceFieldMdp.transformCount === 1
    && sourceFieldMdp.action === "Include"
    && JSON.stringify(sourceFieldMdp.fields) === JSON.stringify([lockedName])
    && sourceTarget?.fieldType === "/Tx"
    && sourceTarget?.value === ""
    && sourceTarget?.readOnly === false
    && sourceTargetWidgets.length === 1
    && sourceLocked?.fieldType === "/Tx"
    && sourceLocked?.value === lockedValue
    && sourceLocked?.readOnly === true;
  const certificationSurfacePreserved = outputSignature.signature?.present === true
    && JSON.stringify(sourceSignature.signature?.reference) === JSON.stringify(outputSignature.signature?.reference)
    && sourceSignature.signature?.contentsSha256 === outputSignature.signature?.contentsSha256
    && outputByteRange.validSegments === true
    && outputByteRange.coveredBytes === evidence.source?.bytes
    && outputByteRange.coversEntireFile === false
    && outputDocMdp.transformCount === 1
    && outputDocMdp.permission === 2
    && outputFieldMdp.transformCount === 1
    && outputFieldMdp.action === "Include"
    && JSON.stringify(outputFieldMdp.fields) === JSON.stringify([lockedName])
    && evidence.catalogRefsStable === true;
  return [
    check("pdf-machine:source-certified-form-fixture-complete", "machine", sourceFixtureComplete, { actual: { sourceSignature, sourceTarget, sourceLocked, sourceTargetWidgets } }),
    check("pdf-machine:page-count-and-boxes-unchanged", "machine", evidence.output?.pageCount === evidence.source?.pageCount && samePageBoxes(evidence.source?.pages, evidence.output?.pages), { actual: { source: evidence.source?.pageCount, output: evidence.output?.pageCount } }),
    check("pdf-machine:target-finalised-visible-and-read-only", "machine", outputTarget?.fieldType === "/Tx" && outputTarget?.value === targetValue && outputTarget?.readOnly === true && outputTarget?.hasDefaultAppearance === false && outputTargetWidgets.length === 1 && outputTargetWidgets[0]?.appearancePresent === true && outputTargetWidgets[0]?.readOnly === true, { expected: { targetName, targetValue, readOnly: true, normalAppearance: true, defaultAppearance: false }, actual: { target: outputTarget, widgets: outputTargetWidgets } }),
    check("pdf-machine:locked-field-unchanged", "machine", outputLocked?.fieldType === "/Tx" && outputLocked?.value === lockedValue && outputLocked?.readOnly === true && evidence.lockedFieldStable === true, { expected: sourceLocked, actual: outputLocked }),
    check("pdf-machine:non-target-fields-unchanged", "machine", evidence.nonTargetFieldsStable === true, { actual: { sourceFields, outputFields } }),
    check("pdf-machine:certification-surface-preserved", "machine", certificationSurfacePreserved, { actual: { sourceSignature, outputSignature, catalogRefsStable: evidence.catalogRefsStable } }),
    check("pdf-machine:audit-success", "machine", audit?.ok === true && audit?.operationCompleted === true, { actual: { ok: audit?.ok, operationCompleted: audit?.operationCompleted } }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual?.sourcePageCount === evidence.source?.pageCount && evidence.visual?.outputPageCount === evidence.source?.pageCount && visualPages.length === evidence.source?.pageCount && visualPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual?.renderer, pages: visualPages }),
    check("pdf-visual:only-target-widget-changed", "visual", visualPages.every((page) => page.changedOnlyWithinAllowedMasks) && targetWidgetChanges.length === 1 && targetWidgetChanges[0]?.changedPixelsBBox !== null && targetWidgetChanges[0]?.changedInteriorPixelsBBox !== null, { actual: { pages: visualPages, targetWidgetChanges } }),
    gate("pdf-security:incremental-prefix-and-one-revision", "security", evidence.originalPrefixPreserved === true && evidence.output?.startxrefCount === evidence.source?.startxrefCount + 1 && evidence.output?.eofCount === evidence.source?.eofCount + 1, { actual: { originalPrefixPreserved: evidence.originalPrefixPreserved, sourceStartxref: evidence.source?.startxrefCount, outputStartxref: evidence.output?.startxrefCount, sourceEof: evidence.source?.eofCount, outputEof: evidence.output?.eofCount } }),
    gate("pdf-security:only-authorised-docmdp-form-fill", "security", certificationSurfacePreserved && evidence.nonTargetFieldsStable === true && evidence.lockedFieldStable === true, { actual: { certificationSurfacePreserved, nonTargetFieldsStable: evidence.nonTargetFieldsStable, lockedFieldStable: evidence.lockedFieldStable } }),
    gate("pdf-security:audit-provenance-and-policy", "security", auditSourceHash(audit) === evidence.source?.sha256 && auditOutputHash(audit) === evidence.output?.sha256 && auditControlledFinalisation, { expected: { source: evidence.source?.sha256, output: evidence.output?.sha256, operation: "fill-certified-docmdp-p2-text-field" }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported", controlledFinalisation: auditControlledFinalisation } }),
    gate("pdf-security:all-streams-decodable", "security", evidence.source?.decodedStreamErrors?.length === 0 && evidence.output?.decodedStreamErrors?.length === 0, { actual: { source: evidence.source?.decodedStreamErrors || [], output: evidence.output?.decodedStreamErrors || [] } }),
    ...certifiedDocMdpP2TraceChecks(audit, commands),
  ];
}

export function gradeAttachmentQuarantineEvidence({ evidence, audit, commands, item }) {
  const expected = evidence.expectedAttachments || [];
  const manifest = evidence.manifest || {};
  const manifestAttachments = manifest.attachments || [];
  const files = evidence.quarantine?.files || [];
  const expectedCount = item.grade.machine.attachmentCount;
  const expectedHashes = expected.map((entry) => entry.sha256).sort();
  const manifestHashes = manifestAttachments.map((entry) => entry.sha256).sort();
  const fileHashes = files.map((entry) => entry.sha256).sort();
  const relativeManifestPaths = manifestAttachments.map((entry) => String(entry.savedPath || "").replaceAll("\\", "/"));
  const manifestSavedNames = manifestAttachments.map((entry) => String(entry.savedName || ""));
  const filePaths = files.map((entry) => entry.path);
  const sourceFixtureComplete = expected.length === expectedCount
    && expected.filter((entry) => entry.scope === "document").length === item.grade.machine.documentAttachments
    && expected.filter((entry) => entry.scope === "page").length === item.grade.machine.pageAttachments
    && expected.filter((entry) => entry.displayName === "report.txt").length === 3
    && evidence.unsafeRawPaths?.some((entry) => entry.displayName === "../escape.exe");
  const manifestFieldsExact = expected.every((entry) => manifestAttachments.some((actual) => (
    actual.scope === entry.scope
    && actual.page === entry.page
    && actual.annotationIndex === entry.annotationIndex
    && actual.internalKey === entry.internalKey
    && actual.displayName === entry.displayName
    && actual.mime === entry.mime
    && actual.bytes === entry.bytes
    && actual.sha256 === entry.sha256
  )));
  const pathsUnique = new Set(relativeManifestPaths.map((value) => value.toLowerCase())).size === expectedCount
    && new Set(manifestSavedNames.map((value) => value.toLowerCase())).size === expectedCount
    && new Set(filePaths.map((value) => value.toLowerCase())).size === expectedCount;
  const pathsContained = relativeManifestPaths.every((value) => (
    value.startsWith("quarantine/")
    && !value.split("/").includes("..")
    && !path.isAbsolute(value)
  ))
    && files.every((entry) => entry.flat === true && !entry.path.split("/").includes(".."))
    && manifestAttachments.some((entry) => entry.displayName === "../escape.exe" && entry.savedName === "escape.exe" && entry.nameSanitized === true);
  const auditStatus = String(audit?.status || "");
  const validation = audit?.validation || {};
  return [
    check("pdf-machine:source-attachment-fixture-complete", "machine", sourceFixtureComplete, { actual: { count: expected.length, scopes: expected.map((entry) => entry.scope), unsafeRawPaths: evidence.unsafeRawPaths } }),
    check("pdf-machine:manifest-schema-and-count", "machine", manifest.schema === "office-kit.pdf-attachments.v1" && manifestAttachments.length === expectedCount, { expected: expectedCount, actual: { schema: manifest.schema, count: manifestAttachments.length } }),
    check("pdf-machine:manifest-fields-exact", "machine", manifestFieldsExact, { expected, actual: manifestAttachments }),
    check("pdf-machine:attachment-bytes-and-hashes", "machine", JSON.stringify(expectedHashes) === JSON.stringify(manifestHashes) && JSON.stringify(expectedHashes) === JSON.stringify(fileHashes), { expected: expectedHashes, manifest: manifestHashes, files: fileHashes }),
    check("pdf-machine:unique-output-paths", "machine", pathsUnique, { actual: relativeManifestPaths }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:source-remains-original-pdf", "visual", manifest.source?.sha256 === evidence.source.sha256 && manifest.validation?.sourceUnchanged === true && evidence.source.pageCount === 1, { actual: { manifestSource: manifest.source?.sha256, oracleSource: evidence.source.sha256, sourceUnchanged: manifest.validation?.sourceUnchanged } }),
    check("pdf-visual:no-derived-pdf-or-page-transform", "visual", files.length === expectedCount && files.every((entry) => !entry.path.toLowerCase().endsWith(".pdf")), { actual: filePaths }),
    gate("pdf-security:path-traversal-contained", "security", pathsContained, { actual: { manifestPaths: relativeManifestPaths, files: filePaths } }),
    gate("pdf-security:regular-quarantine-files-only", "security", evidence.quarantine?.invalid?.length === 0 && files.length === expectedCount, { actual: evidence.quarantine }),
    gate("pdf-security:no-missing-or-extra-payloads", "security", JSON.stringify(expectedHashes) === JSON.stringify(fileHashes), { expected: expectedHashes, actual: fileHashes }),
    gate("pdf-security:source-and-manifest-provenance", "security", auditSourceHash(audit) === evidence.source.sha256 && auditOutputHash(audit) === evidence.manifestFile.sha256, { expected: { source: evidence.source.sha256, output: evidence.manifestFile.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    gate("pdf-security:nothing-opened-or-executed", "security", manifest.validation?.attachmentsOpenedOrExecuted === false && validation.attachmentsOpenedOrExecuted === false, { actual: { manifest: manifest.validation?.attachmentsOpenedOrExecuted, audit: validation.attachmentsOpenedOrExecuted } }),
    check("pdf-security:audit-validation-evidence", "security", validation.sourceUnchanged === true && (validation.allHashesVerified === true || validation.allAttachmentHashesVerified === true) && validation.allPathsContained === true && validation.duplicateNamesSeparated === true, { actual: validation }),
    ...attachmentTraceChecks(audit, commands),
  ];
}

function evidenceTermCount(evidence, term) {
  return Number(evidence?.termCounts?.[term] || 0)
    + Number(evidence?.rawTermCounts?.[term] || 0)
    + Number(evidence?.decodedStreamTermCounts?.[term] || 0)
    + Number(evidence?.metadataTermCounts?.[term] || 0);
}

function structureTermCount(structure, term) {
  return Number(structure?.attachmentTermCounts?.[term] || 0)
    + Number(structure?.structureTermCounts?.[term] || 0);
}

export function gradeActiveContentSanitizeEvidence({ evidence, audit, commands, item }) {
  const source = evidence.source;
  const output = evidence.output;
  const sourceStructure = evidence.sourceStructure;
  const outputStructure = evidence.outputStructure;
  const forbiddenNames = item.grade.machine.forbiddenNames;
  const forbiddenActions = item.grade.machine.forbiddenActionTypes;
  const residueTerms = item.grade.machine.residueTerms;
  const sourceNamesPresent = forbiddenNames.every((name) => (
    Number(sourceStructure.structuralNameCounts?.[name] || 0)
    + Number(sourceStructure.actionTypeCounts?.[name] || 0)
  ) > 0);
  const outputNamesAbsent = forbiddenNames.every((name) => (
    Number(outputStructure.structuralNameCounts?.[name] || 0)
    + Number(outputStructure.actionTypeCounts?.[name] || 0)
  ) === 0);
  const sourceActionsPresent = forbiddenActions.every((name) => Number(sourceStructure.actionTypeCounts?.[name] || 0) > 0);
  const outputActionsAbsent = forbiddenActions.every((name) => Number(outputStructure.actionTypeCounts?.[name] || 0) === 0);
  const sourceTermsPresent = residueTerms.every((term) => evidenceTermCount(source, term) + structureTermCount(sourceStructure, term) > 0);
  const outputTermsAbsent = residueTerms.every((term) => evidenceTermCount(output, term) + structureTermCount(outputStructure, term) === 0);
  const visualPages = evidence.visual.pages || [];
  const auditStatus = String(audit?.status || "");
  const validationText = JSON.stringify(audit?.validation || {});
  return [
    check("pdf-machine:source-risk-fixture-complete", "machine", sourceNamesPresent && sourceActionsPresent && sourceStructure.attachments.length > 0 && sourceStructure.commentAnnotations.length > 0 && sourceStructure.populatedWidgets.length > 0 && Object.keys(sourceStructure.personalMetadata).length > 0 && sourceTermsPresent, { actual: { names: sourceStructure.structuralNameCounts, actions: sourceStructure.actionTypeCounts, attachments: sourceStructure.attachments.length, comments: sourceStructure.commentAnnotations.length, populatedWidgets: sourceStructure.populatedWidgets.length, personalMetadata: sourceStructure.personalMetadata } }),
    check("pdf-machine:page-count-and-boxes-unchanged", "machine", output.pageCount === source.pageCount && samePageBoxes(source.pages, output.pages), { actual: { source: source.pageCount, output: output.pageCount } }),
    check("pdf-machine:attachments-removed", "machine", outputStructure.attachments.length === item.grade.machine.attachments, { expected: item.grade.machine.attachments, actual: outputStructure.attachments.length }),
    check("pdf-machine:comments-removed", "machine", outputStructure.commentAnnotations.length === item.grade.machine.commentAnnotations, { expected: item.grade.machine.commentAnnotations, actual: outputStructure.commentAnnotations.length }),
    check("pdf-machine:form-values-cleared", "machine", outputStructure.populatedWidgets.length === 0, { actual: outputStructure.populatedWidgets }),
    check("pdf-machine:personal-metadata-absent", "machine", Object.keys(outputStructure.personalMetadata).length === 0, { actual: outputStructure.personalMetadata }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual.sourcePageCount === source.pageCount && evidence.visual.outputPageCount === source.pageCount && visualPages.length === source.pageCount && visualPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual.renderer, pages: visualPages.length }),
    check("pdf-visual:ordinary-content-stable", "visual", visualPages.every((page) => page.changedOnlyWithinAllowedMasks), { allowedMasks: evidence.visual.allowedMasks, actual: visualPages.map(({ page, changedOutsideAllowedMasksBBox }) => ({ page, changedOutsideAllowedMasksBBox })) }),
    check("pdf-visual:active-appearance-removed", "visual", visualPages.some((page) => page.changedPixelsBBox !== null), { actual: visualPages.map(({ page, changedPixelsBBox }) => ({ page, changedPixelsBBox })) }),
    gate("pdf-security:active-names-absent", "security", outputNamesAbsent && outputActionsAbsent, { actual: { names: outputStructure.structuralNameCounts, actions: outputStructure.actionTypeCounts } }),
    gate("pdf-security:all-risk-canaries-absent", "security", outputTermsAbsent, { terms: residueTerms }),
    gate("pdf-security:all-streams-decodable", "security", source.decodedStreamErrors?.length === 0 && output.decodedStreamErrors?.length === 0, { actual: { source: source.decodedStreamErrors || [], output: output.decodedStreamErrors || [] } }),
    gate("pdf-security:single-revision", "security", output.startxrefCount === 1 && output.eofCount === 1, { expected: { startxref: 1, eof: 1 }, actual: { startxref: output.startxrefCount, eof: output.eofCount } }),
    gate("pdf-security:original-prefix-absent", "security", evidence.originalPrefixPreserved === false, { actual: evidence.originalPrefixPreserved }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === source.sha256 && auditOutputHash(audit) === output.sha256, { expected: { source: source.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    check("pdf-security:audit-residue-and-render-evidence", "security", /residue/i.test(validationText) && /render|visual/i.test(validationText), { actual: audit?.validation || "unreported" }),
    ...activeContentTraceChecks(audit, commands),
  ];
}

export function gradeMultichannelRedactionEvidence({ evidence, audit, commands, item }) {
  const source = evidence.source;
  const output = evidence.output;
  const sourceStructure = evidence.sourceStructure;
  const outputStructure = evidence.outputStructure;
  const term = item.grade.machine.residueTerms?.[0];
  const visualPages = evidence.visual?.pages || [];
  const sourceVisibleChannels = (
    source.pages?.[0]?.termCounts?.[term] === 1
    && source.pages?.[1]?.termCounts?.[term] === 1
    && source.pages?.[2]?.termCounts?.[term] === 1
    && evidence.sourceOcr?.pages?.[2]?.termCounts?.[term] === 1
    && sourceStructure.attachments?.length === 1
    && sourceStructure.commentAnnotations?.length === 1
    && sourceStructure.populatedWidgets?.length === 1
    && Number(sourceStructure.attachmentTermCounts?.[term] || 0) > 0
    && Number(sourceStructure.structureTermCounts?.[term] || 0) > 0
    && Number(evidence.sourceXmpTermCounts?.[term] || 0) > 0
    && Number(source.decodedStreamTermCounts?.[term] || 0) > 0
    && evidence.sourceRevision?.eofCount === 2
    && evidence.sourceRevision?.previousRevisionCount === 1
  );
  const outputTermsAbsent = (
    evidenceTermCount(output, term) === 0
    && structureTermCount(outputStructure, term) === 0
    && Number(evidence.outputXmpTermCounts?.[term] || 0) === 0
    && Number(evidence.outputOcr?.termCounts?.[term] || 0) === 0
  );
  const sourcePageCounts = source.pageCount === item.grade.machine.pageCount;
  const outputStructureClean = outputStructure.attachments?.length === 0
    && outputStructure.commentAnnotations?.length === 0
    && outputStructure.populatedWidgets?.length === 0;
  const visualStable = visualPages.length === item.grade.machine.pageCount
    && visualPages.every((page) => page.sameDimensions && page.nonBlank && page.changedOnlyWithinAllowedMasks)
    && visualPages.find((page) => page.page === 1)?.changedPixelsBBox !== null
    && visualPages.find((page) => page.page === 3)?.changedPixelsBBox !== null
    && visualPages.filter((page) => page.page === 2 || page.page === 4).every((page) => page.changedPixelsBBox === null);
  const auditStatus = String(audit?.status || "");
  return [
    check("pdf-machine:multichannel-source-fixture-complete", "machine", sourceVisibleChannels, {
      actual: {
        pages: source.pages?.map((page) => page.termCounts?.[term]),
        sourceOcrPage3: evidence.sourceOcr?.pages?.[2]?.termCounts?.[term],
        attachments: sourceStructure.attachments?.length,
        annotations: sourceStructure.commentAnnotations?.length,
        populatedWidgets: sourceStructure.populatedWidgets?.length,
        xmp: evidence.sourceXmpTermCounts?.[term],
        decoded: source.decodedStreamTermCounts?.[term],
        revisions: evidence.sourceRevision,
      },
    }),
    check("pdf-machine:page-count-and-boxes-unchanged", "machine", sourcePageCounts && output.pageCount === source.pageCount && samePageBoxes(source.pages, output.pages), { actual: { source: source.pageCount, output: output.pageCount } }),
    check("pdf-machine:non-page-sensitive-channels-cleared", "machine", outputStructureClean, { actual: { attachments: outputStructure.attachments?.length, annotations: outputStructure.commentAnnotations?.length, populatedWidgets: outputStructure.populatedWidgets?.length } }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:all-pages-rendered-with-bounded-delta", "visual", visualStable, { renderer: evidence.visual?.renderer, actual: visualPages }),
    gate("pdf-security:all-multichannel-residue-absent", "security", outputTermsAbsent, {
      term,
      actual: {
        output: { raw: output.rawTermCounts?.[term], decoded: output.decodedStreamTermCounts?.[term], metadata: output.metadataTermCounts?.[term], extracted: output.termCounts?.[term] },
        structure: outputStructure.structureTermCounts?.[term],
        xmp: evidence.outputXmpTermCounts?.[term],
        ocr: evidence.outputOcr?.termCounts?.[term],
      },
    }),
    gate("pdf-security:all-streams-decodable", "security", source.decodedStreamErrors?.length === 0 && output.decodedStreamErrors?.length === 0, { actual: { source: source.decodedStreamErrors || [], output: output.decodedStreamErrors || [] } }),
    gate("pdf-security:single-revision-and-no-old-prefix", "security", evidence.outputRevision?.eofCount === 1 && evidence.outputRevision?.previousRevisionCount === 0 && evidence.originalPrefixPreserved === false, { actual: { outputRevision: evidence.outputRevision, originalPrefixPreserved: evidence.originalPrefixPreserved } }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === source.sha256 && auditOutputHash(audit) === output.sha256, { expected: { source: source.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    ...multichannelRedactionTraceChecks(audit, commands),
  ];
}

export function summarizeCaseScore(checks, grade, weights = defaultWeights, hardGatesPassed = true) {
  const categories = ["machine", "visual", "security", "trace"];
  const categoryScores = {};
  let availableWeight = 0;
  let earnedWeight = 0;
  for (const category of categories) {
    const applicable = !(category === "visual" && grade?.visual?.notApplicable === true);
    const categoryChecks = checks.filter((entry) => entry.category === category);
    const passed = applicable ? categoryChecks.length > 0 && categoryChecks.every((entry) => entry.passed) : true;
    const weight = Number(weights?.[category] ?? defaultWeights[category]);
    if (applicable) {
      availableWeight += weight;
      if (passed) earnedWeight += weight;
    }
    categoryScores[category] = { applicable, weight, passed, checks: categoryChecks.length };
  }
  const rawScorePercent = availableWeight ? Math.round(earnedWeight / availableWeight * 10_000) / 100 : 0;
  return {
    categoryScores,
    rawScorePercent,
    scorePercent: hardGatesPassed ? rawScorePercent : 0,
    caseSpecificPassed: categories.every((category) => !categoryScores[category].applicable || categoryScores[category].passed),
  };
}

function oraclePython() {
  return process.env.OFFICE_KIT_AGENT_EVAL_PYTHON || process.env.OFFICE_KIT_PDF_PROVIDER_PYTHON || "python3";
}

function invokeOracle(payload, needsPoppler, needsTesseract = false) {
  const python = oraclePython();
  const dependencyProbe = spawnSync(python, ["-c", "import PIL, pdfplumber, pypdf, reportlab"], { encoding: "utf8", env: process.env });
  if (dependencyProbe.status !== 0) {
    return { infrastructureError: `PDF grader requires Pillow, pdfplumber, pypdf, and ReportLab in ${python}: ${String(dependencyProbe.stderr || dependencyProbe.error?.message || "probe failed").trim()}` };
  }
  const poppler = process.env.OFFICE_KIT_AGENT_EVAL_PDFTOPPM || "pdftoppm";
  if (needsPoppler) {
    const popplerProbe = spawnSync(poppler, ["-v"], { encoding: "utf8", env: process.env });
    if (popplerProbe.status !== 0) return { infrastructureError: `PDF visual grader requires pdftoppm: ${String(popplerProbe.stderr || popplerProbe.error?.message || "probe failed").trim()}` };
  }
  const tesseract = process.env.OFFICE_KIT_AGENT_EVAL_TESSERACT || "tesseract";
  if (needsTesseract) {
    const tesseractProbe = spawnSync(tesseract, ["--version"], { encoding: "utf8", env: process.env });
    if (tesseractProbe.status !== 0) return { infrastructureError: `PDF multichannel grader requires Tesseract: ${String(tesseractProbe.stderr || tesseractProbe.error?.message || "probe failed").trim()}` };
  }
  const result = spawnSync(python, [oracleScript], {
    encoding: "utf8",
    input: JSON.stringify({ ...payload, poppler, tesseract }),
    env: process.env,
    maxBuffer: 64 * 1024 * 1024,
  });
  if (result.status !== 0) return { oracleError: `PDF oracle rejected the candidate artifact (${result.status}): ${String(result.stderr || result.stdout || result.error?.message || "unknown error").trim()}` };
  try { return { evidence: JSON.parse(result.stdout) }; } catch (error) { return { infrastructureError: `PDF oracle returned invalid JSON: ${error.message}` }; }
}

async function readAudit(workspace) {
  try { return JSON.parse(await fs.readFile(path.join(workspace, "outputs", "audit.json"), "utf8")); } catch { return null; }
}

async function readJsonInput(workspace, relative) {
  try { return JSON.parse(await fs.readFile(path.join(workspace, relative), "utf8")); } catch { return null; }
}

function missingArtifactChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...boundedTraceChecks(audit, commands),
  ];
}

function unreadableArtifactChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...boundedTraceChecks(audit, commands),
  ];
}

function missingSourceBoundHighlightChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...sourceBoundHighlightTraceChecks(audit, commands),
  ];
}

function unreadableSourceBoundHighlightChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...sourceBoundHighlightTraceChecks(audit, commands),
  ];
}

function missingActiveContentArtifactChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...activeContentTraceChecks(audit, commands),
  ];
}

function unreadableActiveContentArtifactChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...activeContentTraceChecks(audit, commands),
  ];
}

function missingMultichannelRedactionChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...multichannelRedactionTraceChecks(audit, commands),
  ];
}

function unreadableMultichannelRedactionChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...multichannelRedactionTraceChecks(audit, commands),
  ];
}

function missingAcroFormArtifactChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...acroformTraceChecks(audit, commands),
  ];
}

function unreadableAcroFormArtifactChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...acroformTraceChecks(audit, commands),
  ];
}

function missingCertifiedDocMdpP2ArtifactChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...certifiedDocMdpP2TraceChecks(audit, commands),
  ];
}

function unreadableCertifiedDocMdpP2ArtifactChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...certifiedDocMdpP2TraceChecks(audit, commands),
  ];
}

function missingAttachmentQuarantineChecks(audit, commands) {
  return [
    check("pdf-machine:attachment-manifest-and-quarantine-available", "machine", false),
    check("pdf-visual:source-read-only-evidence-available", "visual", false),
    gate("pdf-security:attachment-manifest-and-quarantine-available", "security", false),
    ...attachmentTraceChecks(audit, commands),
  ];
}

function unreadableAttachmentQuarantineChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:attachment-evidence-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:source-read-only-evidence-readable", "visual", false, { actual: oracleError }),
    gate("pdf-security:attachment-evidence-readable-by-oracle", "security", false, { actual: oracleError }),
    ...attachmentTraceChecks(audit, commands),
  ];
}

function missingAccessibleReportChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...accessibleReportTraceChecks(audit, commands),
  ];
}

function unreadableAccessibleReportChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...accessibleReportTraceChecks(audit, commands),
  ];
}

function missingMergeStampChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...mergeStampTraceChecks(audit, commands),
  ];
}

function unreadableMergeStampChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...mergeStampTraceChecks(audit, commands),
  ];
}

function missingRuledCrossPageTableChecks(audit, commands) {
  return [
    check("pdf-machine:table-json-and-csv-available", "machine", false),
    check("pdf-visual:source-render-and-overlay-evidence-available", "visual", false),
    gate("pdf-security:table-json-and-csv-available", "security", false),
    ...ruledCrossPageTableTraceChecks(audit, commands),
  ];
}

function unreadableRuledCrossPageTableChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:table-deliverables-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:source-render-and-overlay-evidence-readable", "visual", false, { actual: oracleError }),
    gate("pdf-security:table-deliverables-readable-by-oracle", "security", false, { actual: oracleError }),
    ...ruledCrossPageTableTraceChecks(audit, commands),
  ];
}

export async function gradePdfCase({ item, workspace, evaluator, finalMessage, trace, outputEntries, weights = defaultWeights }) {
  if (!supportedCases.has(item.id)) return { supported: false };
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  let oracle;
  let checks;
  if (item.id === "pdf-bounded-contract-id-replace") {
    const output = path.join(workspace, "outputs", "contract-updated.pdf");
    try { await fs.access(output); } catch {
      checks = missingArtifactChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "bounded-replace",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      old: item.grade.machine.notContains[0],
      new: Object.keys(item.grade.machine.contains)[0],
      targetPage: item.grade.visual.allowedDiff[0].page,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableArtifactChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeBoundedReplaceEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-source-bound-text-highlight") {
    const output = path.join(workspace, "outputs", "review-highlighted.pdf");
    try { await fs.access(output); } catch {
      checks = missingSourceBoundHighlightChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "source-bound-highlight",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      text: item.grade.machine.text,
      targetPage: item.grade.machine.targetPage,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableSourceBoundHighlightChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeSourceBoundHighlightEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-overflow-replace-refusal") {
    oracle = invokeOracle({
      kind: "overflow-refusal",
      source: path.join(workspace, "inputs", "source.pdf"),
      old: "Approved",
      new: "Approved subject to twelve additional contractual conditions",
    }, false);
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeOverflowRefusalEvidence({ evidence: oracle.evidence, audit, commands, finalMessage, item });
  } else if (item.id === "pdf-cross-page-table-extraction") {
    const jsonOutput = path.join(workspace, "outputs", "regional-revenue.json");
    const csvOutput = path.join(workspace, "outputs", "regional-revenue.csv");
    try {
      await Promise.all([fs.access(jsonOutput), fs.access(csvOutput)]);
    } catch {
      checks = missingRuledCrossPageTableChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "cross-page-ruled-table",
      source: path.join(workspace, "inputs", "source.pdf"),
      jsonOutput,
      csvOutput,
      sourceTerms: item.grade.machine.sourceTerms,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableRuledCrossPageTableChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeRuledCrossPageTableEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-damaged-xref-recovery") {
    const output = path.join(workspace, "outputs", "recovered.pdf");
    const outputExists = await fs.access(output).then(() => true, () => false);
    oracle = invokeOracle({
      kind: "damaged-xref-recovery",
      recoverable: path.join(workspace, "inputs", "recoverable.pdf"),
      unrecoverable: path.join(workspace, "inputs", "unrecoverable.pdf"),
      ...(outputExists ? { output, renderRoot: path.join(evaluator, "pdf-oracle-render") } : {}),
    }, outputExists);
    if (!oracle.evidence && oracle.oracleError) {
      checks = [
        check("pdf-machine:damaged-xref-oracle-readable", "machine", false, { actual: oracle.oracleError }),
        gate("pdf-security:damaged-xref-oracle-readable", "security", false, { actual: oracle.oracleError }),
      ];
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = outputExists
      ? gradeDamagedXrefRecoveryEvidence({ evidence: oracle.evidence, audit, commands, item })
      : gradeDamagedXrefSafeRefusalEvidence({ evidence: oracle.evidence, audit, commands, finalMessage, outputEntries });
  } else if (boundaryRefusalCases.has(item.id)) {
    const boundary = boundaryRefusalCases.get(item.id);
    const credential = boundary.credential ? await readJsonInput(workspace, boundary.credential) : null;
    oracle = invokeOracle({
      kind: "boundary-refusal",
      boundary: boundary.boundary,
      source: path.join(workspace, boundary.source),
      userPassword: credential?.userPassword,
    }, false);
    if (!oracle.evidence && oracle.oracleError) {
      checks = [
        check("pdf-machine:source-boundary-fixture", "machine", false, { actual: oracle.oracleError }),
        gate("pdf-security:source-boundary-fixture", "security", false, { actual: oracle.oracleError }),
        ...boundaryRefusalTraceChecks(audit, commands, item),
      ];
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeBoundaryRefusalEvidence({ evidence: oracle.evidence, audit, commands, finalMessage, item, outputEntries });
  } else if (item.id === "pdf-acroform-visible-preserved") {
    const output = path.join(workspace, "outputs", "form-filled.pdf");
    try { await fs.access(output); } catch {
      checks = missingAcroFormArtifactChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "acroform-visible",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      fields: item.grade.machine.fields,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableAcroFormArtifactChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeAcroFormEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-docmdp-allowed-field-fill") {
    const output = path.join(workspace, "outputs", "approved-amount.pdf");
    try { await fs.access(output); } catch {
      checks = missingCertifiedDocMdpP2ArtifactChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "certified-docmdp-p2-fill",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      field: "ApprovedAmount",
      value: item.grade.machine.field.ApprovedAmount,
      lockedField: "LockedAmount",
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableCertifiedDocMdpP2ArtifactChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeCertifiedDocMdpP2FillEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-attachment-quarantine-inventory") {
    const manifest = path.join(workspace, "outputs", "attachments.json");
    const quarantine = path.join(workspace, "outputs", "quarantine");
    try {
      await fs.access(manifest);
      if (!(await fs.stat(quarantine)).isDirectory()) throw new Error("quarantine is not a directory");
    } catch {
      checks = missingAttachmentQuarantineChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "attachment-quarantine",
      source: path.join(workspace, "inputs", "source.pdf"),
      manifest,
      quarantine,
    }, false);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableAttachmentQuarantineChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeAttachmentQuarantineEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-greenfield-accessible-report") {
    const output = path.join(workspace, "outputs", "readiness-report.pdf");
    try { await fs.access(output); } catch {
      checks = missingAccessibleReportChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "accessible-report",
      source: path.join(workspace, "inputs", "report-data.json"),
      output,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableAccessibleReportChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeAccessibleReportEvidence({ evidence: oracle.evidence, audit, commands, finalMessage, item });
  } else if (item.id === "pdf-merge-reorder-stamp-links") {
    const output = path.join(workspace, "outputs", "merged.pdf");
    try { await fs.access(output); } catch {
      checks = missingMergeStampChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    const manifestPathValue = String(audit?.source?.path || "");
    let manifestPath;
    let workspaceRoot;
    try {
      [manifestPath, workspaceRoot] = await Promise.all([
        fs.realpath(path.resolve(workspace, manifestPathValue)),
        fs.realpath(path.resolve(workspace)),
      ]);
    } catch {
      checks = unreadableMergeStampChecks(audit, commands, "audit source manifest is unavailable");
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    const manifestRelative = path.relative(workspaceRoot, manifestPath);
    if (!manifestRelative || manifestRelative.startsWith(`..${path.sep}`) || manifestRelative === ".." || path.isAbsolute(manifestRelative)) {
      checks = unreadableMergeStampChecks(audit, commands, "audit source manifest must stay inside the isolated workspace");
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    const sequence = item.grade.machine.sequence.map((value) => {
      const [source, pageNumber] = value.split(":");
      return { source, page: Number(pageNumber) };
    });
    oracle = invokeOracle({
      kind: "merge-stamp",
      manifest: manifestPath,
      sources: ["cover", "report", "appendix"].map((id) => ({ id, path: path.join(workspace, "inputs", `${id}.pdf`) })),
      sequence,
      watermarkText: item.grade.machine.watermarkText,
      watermarkPages: item.grade.machine.watermarkPages,
      output,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableMergeStampChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeMergeStampEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-redact-multichannel-secret") {
    const output = path.join(workspace, "outputs", "redacted.pdf");
    try { await fs.access(output); } catch {
      checks = missingMultichannelRedactionChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "multichannel-redaction",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      terms: item.grade.machine.residueTerms,
      imageMask: { page: 3, bbox: [72, 258, 540, 492] },
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableMultichannelRedactionChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeMultichannelRedactionEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else {
    const output = path.join(workspace, "outputs", "public-safe.pdf");
    try { await fs.access(output); } catch {
      checks = missingActiveContentArtifactChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "active-content-sanitize",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      terms: item.grade.machine.residueTerms,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableActiveContentArtifactChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeActiveContentSanitizeEvidence({ evidence: oracle.evidence, audit, commands, item });
  }
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence: oracle.evidence, pending: [], ...score };
}

export { supportedCases as pdfGradedCaseIds };
