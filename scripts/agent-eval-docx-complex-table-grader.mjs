import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import { DOCX_COMPLEX_TABLE_TOPOLOGY_BOUNDARY_FIXTURE } from "./agent-eval-office-fixtures.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };

function check(id, category, passed, details = {}) {
  return { id, category, gate: false, passed: Boolean(passed), ...details };
}

function gate(id, category, passed, details = {}) {
  return { id, category, gate: true, passed: Boolean(passed), ...details };
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function decodeXml(value = "") {
  return String(value)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", "\"")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function xmlAttributes(opening = "") {
  const result = {};
  for (const match of String(opening).matchAll(/([:\w.-]+)="([^"]*)"/g)) {
    result[match[1].split(":").at(-1)] = decodeXml(match[2]);
  }
  return result;
}

function wordText(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
    .map((match) => decodeXml(match[1].replace(/<[^>]+>/g, "")))
    .join("");
}

function balancedWordTables(xml = "") {
  const tables = [];
  const stack = [];
  const tags = /<(\/?)w:tbl\b[^>]*>/g;
  for (const match of String(xml).matchAll(tags)) {
    if (!match[1]) {
      stack.push({ start: match.index, opening: match[0], depth: stack.length + 1 });
      continue;
    }
    const current = stack.pop();
    if (!current) return { balanced: false, tables: [] };
    tables.push({
      ...current,
      end: match.index + match[0].length,
      xml: String(xml).slice(current.start, match.index + match[0].length),
    });
  }
  return { balanced: stack.length === 0, tables: tables.sort((left, right) => left.start - right.start) };
}

function outerTableProperties(tableXml = "") {
  const match = /^<w:tbl\b[^>]*>\s*<w:tblPr\b[^>]*>([\s\S]*?)<\/w:tblPr>/i.exec(String(tableXml));
  return match?.[1] || "";
}

function firstTagAttributes(xml = "", localName) {
  const match = new RegExp(`<(?:[\\w.-]+:)?${localName}\\b[^>]*\\/?\\s*>`, "i").exec(String(xml));
  return match ? xmlAttributes(match[0]) : {};
}

function partHashMap(zip, paths) {
  return Promise.all(paths.map(async (partPath) => [partPath, sha256(await zip.file(partPath)?.async("uint8array") || [])]));
}

/**
 * Reads the source package directly. It intentionally does not call
 * DocumentFile: the public projection is expected to leave this topology
 * source-bound, so the evaluator owns the native-profile assertion.
 */
export async function inspectDocxComplexTableTopology(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const documentXml = await zip.file("word/document.xml")?.async("text") || "";
  const stylesXml = await zip.file("word/styles.xml")?.async("text") || "";
  const parsedTables = balancedWordTables(documentXml);
  const tableRecords = parsedTables.tables.map((table) => {
    const properties = outerTableProperties(table.xml);
    const caption = firstTagAttributes(properties, "tblCaption").val || "";
    const styleId = firstTagAttributes(properties, "tblStyle").val || "";
    return {
      depth: table.depth,
      start: table.start,
      end: table.end,
      caption,
      styleId,
      text: wordText(table.xml),
      xml: table.xml,
    };
  });
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: Object.fromEntries(await partHashMap(zip, paths)),
    documentXml,
    stylesXml,
    tablesBalanced: parsedTables.balanced,
    tables: tableRecords,
  };
}

export function complexTableTopologyProfile(source) {
  const fixture = DOCX_COMPLEX_TABLE_TOPOLOGY_BOUNDARY_FIXTURE;
  const topLevelTables = source.tables.filter((table) => table.depth === 1);
  const complexTable = topLevelTables.find((table) => table.caption === fixture.tableCaption) || null;
  const nestedTables = complexTable
    ? source.tables.filter((table) => table.start > complexTable.start && table.end < complexTable.end)
    : [];
  const directNestedTables = complexTable
    ? nestedTables.filter((table) => table.depth === complexTable.depth + 1)
    : [];
  const properties = outerTableProperties(complexTable?.xml || "");
  const customStyles = [...String(source.stylesXml).matchAll(/<w:style\b[^>]*>/gi)]
    .filter((match) => {
      const attributes = xmlAttributes(match[0]);
      return attributes.type === "table" && attributes.styleId === fixture.styleId;
    });
  const restartMerges = [...String(complexTable?.xml || "").matchAll(/<w:vMerge\b(?=[^>]*\bw:val="restart")[^>]*\/>/gi)];
  const continuationMerges = [...String(complexTable?.xml || "").matchAll(/<w:vMerge\b(?![^>]*\bw:val=)[^>]*\/>/gi)];
  const taggedControls = [...String(complexTable?.xml || "").matchAll(/<w:sdt\b[\s\S]*?<w:tag\b[^>]*\bw:val="ROUTE_CONTROL"[^>]*\/>[\s\S]*?<\/w:sdt>/gi)];
  const matchingRevisions = [...String(complexTable?.xml || "").matchAll(/<w:ins\b(?=[^>]*\bw:id="37")(?=[^>]*\bw:author="Clinical QA")(?=[^>]*\bw:date="2026-07-28T08:00:00Z")[^>]*>/gi)];
  const style = customStyles.length === 1;
  const verticalMerge = restartMerges.length === 1 && continuationMerges.length === 1;
  const contentControl = taggedControls.length === 1;
  const revision = matchingRevisions.length === 1 && wordText(complexTable?.xml || "").includes(fixture.revision.text);
  const headers = fixture.headers.every((header) => complexTable?.text.includes(header));
  const bodyCanaries = complexTable?.text.includes(fixture.mergedDose)
    && complexTable?.text.includes(fixture.routeValue)
    && complexTable?.text.includes(fixture.nestedScheduleLabel)
    && fixture.nestedSchedule.every((text) => complexTable?.text.includes(text));
  const outerStyle = firstTagAttributes(properties, "tblStyle").val === fixture.styleId;
  return {
    ok: source.tablesBalanced
      && topLevelTables.length === 2
      && Boolean(complexTable)
      && outerStyle
      && style
      && verticalMerge
      && nestedTables.length === 1
      && directNestedTables.length === 1
      && contentControl
      && revision
      && headers
      && bodyCanaries,
    topLevelTableCount: topLevelTables.length,
    complexTableFound: Boolean(complexTable),
    outerStyle,
    customStyle: style,
    customStyleCount: customStyles.length,
    verticalMerge,
    verticalMergeRestartCount: restartMerges.length,
    verticalMergeContinuationCount: continuationMerges.length,
    nestedTableCount: nestedTables.length,
    directNestedTableCount: directNestedTables.length,
    contentControl,
    contentControlCount: taggedControls.length,
    revision,
    revisionCount: matchingRevisions.length,
    headers,
    bodyCanaries,
  };
}

function auditHash(audit, side) {
  const record = audit?.[side] || {};
  return String(record.sha256 || audit?.[`${side}Sha256`] || audit?.[`${side}_sha256`] || "");
}

function auditProvider(audit) {
  const provider = audit?.provider;
  if (typeof provider === "string") return provider;
  const officeKit = audit?.officeKit || audit?.officekit || {};
  const actualProvider = audit?.actualProvider || audit?.actual_provider || {};
  if (typeof actualProvider === "string") return actualProvider;
  return [provider?.actual, provider?.selected, provider?.name, provider?.package, provider?.provider, officeKit.actualProvider, officeKit.actual_provider, actualProvider.name, actualProvider.package, actualProvider.provider]
    .find((candidate) => typeof candidate === "string" && candidate.trim()) || "";
}

function auditProviderVersion(audit) {
  const provider = audit?.provider || {};
  const officeKit = audit?.officeKit || audit?.officekit || {};
  const actualProvider = audit?.actualProvider || audit?.actual_provider || {};
  return [provider.version, provider.actualVersion, provider.actual_version, audit?.providerVersion, audit?.provider_version, officeKit.actualVersion, officeKit.actual_version, actualProvider.version, actualProvider.actualVersion, actualProvider.actual_version]
    .find((candidate) => typeof candidate === "string" && candidate.trim()) || "";
}

function auditFallbackIsFalse(audit) {
  const provider = audit?.provider || {};
  const officeKit = audit?.officeKit || audit?.officekit || {};
  const values = [
    provider.silentFallback,
    provider.silent_fallback,
    provider.fallbackUsed,
    provider.fallback_used,
    audit?.silentFallback,
    audit?.silent_fallback,
    audit?.fallbackUsed,
    audit?.fallback_used,
    officeKit.silentFallback,
    officeKit.silent_fallback,
    officeKit.fallbackUsed,
    officeKit.fallback_used,
  ].filter((value) => value !== undefined);
  return values.length > 0 && values.every((value) => value === false || value === "false");
}

function typedComplexTablePreflight(audit) {
  const fixture = DOCX_COMPLEX_TABLE_TOPOLOGY_BOUNDARY_FIXTURE;
  const preflight = audit?.preflight || audit?.validation?.preflight || {};
  const importRecord = preflight.import || audit?.validation?.import || {};
  const inspectRecord = preflight.inspect || audit?.validation?.inspect || {};
  const verifyRecord = preflight.verify || audit?.validation?.verify || {};
  const profile = inspectRecord.profile || inspectRecord.complexTableProfile || audit?.validation?.complexTableProfile || {};
  const operation = audit?.operation || {};
  const target = operation.target || audit?.target || {};
  const targetBound = target.tableCaption === fixture.tableCaption
    || target.caption === fixture.tableCaption
    || target === fixture.tableCaption;
  const sourceBound = preflight.capabilityDecision?.supported === false
    || audit?.validation?.capabilityDecision?.supported === false
    || audit?.capability?.complexTableTopologyEdit === false;
  const profileMatch = inspectRecord.tableCaption === fixture.tableCaption
    && inspectRecord.topLevelTableCount === 2
    && inspectRecord.nestedTableCount === 1
    && inspectRecord.verticalMerge === true
    && inspectRecord.customStyle === true
    && inspectRecord.revision === true
    && inspectRecord.contentControl === true
    || profile.tableCaption === fixture.tableCaption
      && profile.topLevelTableCount === 2
      && profile.nestedTableCount === 1
      && profile.verticalMerge === true
      && profile.customStyle === true
      && profile.revision === true
      && profile.contentControl === true;
  const noMutation = operation.mutationAttempted === false
    || operation.attempted === false
    || audit?.mutationAttempted === false;
  return importRecord.ok === true
    && /(?:DocumentFile\.)?importDocx/i.test(String(importRecord.method || importRecord.operation || ""))
    && inspectRecord.ok === true
    && profileMatch
    && sourceBound
    && noMutation
    && verifyRecord.ok === true
    && targetBound;
}

function hasMutationBypass(commandText) {
  return [
    /(?:DocumentFile\.)?exportDocx/i,
    /(?:DocumentFile\.)?patchDocx/i,
    /\b(?:JSZip|AdmZip|zipfile|python-docx)\b/i,
    /(?:word\/document\.xml|word\/styles\.xml)[^\n]{0,240}(?:write|replace|append|set|patch)/i,
    /\b(?:unzip|zip)\b/i,
    /(?:addColumn|insertColumn|table\.columns)/i,
  ].some((pattern) => pattern.test(commandText));
}

function failedClosedAudit(audit, source) {
  if (!audit) return { present: false, valid: true };
  const output = audit.output;
  const noArtifact = output === null || output === undefined
    || output?.path === null || output?.path === undefined;
  return {
    present: true,
    valid: audit?.status === "failed_closed"
      && auditHash(audit, "source") === source.sha256
      && noArtifact
      && audit?.deliveredModifiedDocx !== true
      && audit?.delivered_modified_docx !== true
      && /office[- ]?kit/i.test(auditProvider(audit))
      && Boolean(auditProviderVersion(audit))
      && auditFallbackIsFalse(audit),
  };
}

function hasTypedInspection(commandText, audit) {
  if (!typedComplexTablePreflight(audit)) return false;
  return /(?:DocumentFile\.)?importDocx/i.test(commandText)
    || /(?:^|\/)officekit(?:\.mjs)?\s+run\b|node_modules\/.bin\/officekit\s+run\b/i.test(commandText)
    || /(?:^|[^\w.-])(?:node|nodejs)\b[^\n]{0,320}?\.(?:mjs|js)(?=\s|['"]|$)/i.test(commandText);
}

export function gradeDocxComplexTableTopologyBoundaryEvidence({ evidence, audit, commands, item }) {
  const profile = complexTableTopologyProfile(evidence.source);
  const commandText = commands.join("\n");
  const diagnostics = item.grade?.machine?.diagnosticTerms || [];
  const diagnosticText = `${evidence.finalMessage || ""}\n${audit?.reason || ""}\n${JSON.stringify(audit?.warnings || [])}`.toLowerCase();
  const auditResult = failedClosedAudit(audit, evidence.source);
  const mutationBypass = hasMutationBypass(commandText);
  return [
    check("docx-complex-table-machine:fixture-profile", "machine", profile.ok, { profile }),
    check("docx-complex-table-machine:explicit-safe-refusal", "machine", diagnostics.some((term) => diagnosticText.includes(String(term).toLowerCase())), {
      expectedAny: diagnostics,
      actual: diagnosticText,
    }),
    check("docx-complex-table-machine:audit-is-failed-closed-when-present", "machine", auditResult.valid, {
      auditPresent: auditResult.present,
      audit: audit || null,
    }),
    gate("docx-complex-table-security:merged-nested-styled-revised-sdt-profile", "security", profile.ok, { profile }),
    gate("docx-complex-table-security:no-output-or-package-mutation", "security", !mutationBypass, {
      forbidden: "DOCX export/patch, direct ZIP access, or table-topology mutation after the source-bound refusal",
      commands,
    }),
    check("docx-complex-table-trace:typed-import-inspect-preflight", "trace", hasTypedInspection(commandText, audit) && typedComplexTablePreflight(audit), {
      expected: "public DocumentFile.importDocx plus a source-bound complex-table preflight audit",
      commands,
    }),
    gate("docx-complex-table-trace:no-silent-fallback", "trace", !mutationBypass && auditFallbackIsFalse(audit), {
      provider: audit?.provider || null,
      commands,
    }),
  ];
}

async function readAudit(workspace) {
  try {
    return JSON.parse(await fs.readFile(path.join(workspace, "outputs", "audit.json"), "utf8"));
  } catch {
    return null;
  }
}

export async function gradeDocxComplexTableTopologyBoundaryCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const fixture = DOCX_COMPLEX_TABLE_TOPOLOGY_BOUNDARY_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  let source;
  try {
    source = await inspectDocxComplexTableTopology(path.join(workspace, "inputs", fixture.documentName));
  } catch (error) {
    const checks = [
      gate("docx-complex-table-machine:readable-source", "machine", false, { error: error.message }),
      gate("docx-complex-table-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }
  const evidence = { source, finalMessage };
  const checks = gradeDocxComplexTableTopologyBoundaryEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}
