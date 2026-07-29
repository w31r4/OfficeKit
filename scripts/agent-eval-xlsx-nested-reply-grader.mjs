import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import { XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE } from "./agent-eval-office-fixtures.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const THREADED_COMMENTS_CONTENT_TYPE = "application/vnd.ms-excel.threadedcomments+xml";
const PERSON_CONTENT_TYPE = "application/vnd.ms-excel.person+xml";
const THREADED_COMMENTS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment";
const PERSON_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2017/10/relationships/person";

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

function innerText(xml = "") {
  return decodeXml(String(xml).replace(/<[^>]+>/g, ""));
}

function normalizeIso(value) {
  const milliseconds = Date.parse(String(value || ""));
  return Number.isFinite(milliseconds) ? new Date(milliseconds).toISOString() : null;
}

function parseThreadedComments(xml = "") {
  const comments = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?threadedComment\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?threadedComment>/g)) {
    const opening = /^<(?:[\w.-]+:)?threadedComment\b[^>]*>/.exec(match[0])?.[0] || "";
    const attributes = xmlAttributes(opening);
    comments.push({
      id: String(attributes.id || "").toUpperCase(),
      parentId: attributes.parentId ? String(attributes.parentId).toUpperCase() : null,
      personId: String(attributes.personId || "").toUpperCase(),
      ref: String(attributes.ref || "").toUpperCase(),
      date: normalizeIso(attributes.dT),
      done: new Set(["1", "true", "on"]).has(String(attributes.done || "0").toLowerCase()),
      text: innerText(match[1]),
    });
  }
  return comments;
}

function parsePeople(xml = "") {
  const people = new Map();
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?person\b[^>]*\/?\s*>/g)) {
    const attributes = xmlAttributes(match[0]);
    const id = String(attributes.id || "").toUpperCase();
    if (!id) continue;
    people.set(id, {
      id,
      displayName: attributes.displayName || "",
      userId: attributes.userId || "",
      providerId: attributes.providerId || "",
    });
  }
  return people;
}

function relationshipSource(partPath) {
  if (partPath === "_rels/.rels") return "";
  const match = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/.exec(partPath);
  if (!match) return null;
  return match[1] ? `${match[1]}/${match[2]}` : match[2];
}

function resolveRelationshipTarget(source, target) {
  const value = String(target || "").replaceAll("\\", "/").split("#")[0];
  if (value.startsWith("/")) return value.slice(1);
  const directory = source ? path.posix.dirname(source) : "";
  return path.posix.normalize(path.posix.join(directory === "." ? "" : directory, value));
}

function parseRelationships(partPath, xml) {
  const source = relationshipSource(partPath);
  if (source === null) return [];
  const relationships = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?Relationship\b[^>]*\/?\s*>/g)) {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type || !attributes.Target) continue;
    relationships.push({
      source,
      relationshipPath: partPath,
      id: attributes.Id,
      type: attributes.Type,
      target: resolveRelationshipTarget(source, attributes.Target),
      external: String(attributes.TargetMode || "").toLowerCase() === "external",
    });
  }
  return relationships;
}

function parseContentTypes(xml = "") {
  const overrides = new Map();
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?Override\b[^>]*\/?\s*>/g)) {
    const attributes = xmlAttributes(match[0]);
    const partName = String(attributes.PartName || "").replace(/^\/+/, "");
    if (partName && attributes.ContentType) overrides.set(partName, attributes.ContentType);
  }
  return overrides;
}

async function packagePartHashes(zip, paths) {
  const hashes = {};
  for (const partPath of paths) hashes[partPath] = sha256(await zip.file(partPath).async("uint8array"));
  return hashes;
}

export async function inspectXlsxNestedThreadedReplyGraph(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const threadedPaths = paths.filter((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name));
  const personPaths = paths.filter((name) => /^xl\/persons\/[^/]+\.xml$/i.test(name));
  const classicCommentPaths = paths.filter((name) => /^xl\/comments\d*\.xml$/i.test(name));
  const [threadedXml, personXml, contentTypesXml, workbookXml, relationshipPayloads] = await Promise.all([
    Promise.all(threadedPaths.map((partPath) => zip.file(partPath).async("text"))),
    Promise.all(personPaths.map((partPath) => zip.file(partPath).async("text"))),
    zip.file("[Content_Types].xml")?.async("text") || "",
    zip.file("xl/workbook.xml")?.async("text") || "",
    Promise.all(paths.filter((partPath) => partPath.endsWith(".rels")).map(async (partPath) => [partPath, await zip.file(partPath).async("text")])),
  ]);
  const people = new Map();
  for (const xml of personXml) for (const [id, person] of parsePeople(xml)) people.set(id, person);
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: await packagePartHashes(zip, paths),
    threadedPaths,
    personPaths,
    classicCommentPaths,
    comments: threadedXml.flatMap(parseThreadedComments),
    people: [...people.values()],
    contentTypes: parseContentTypes(contentTypesXml),
    relationships: relationshipPayloads.flatMap(([partPath, xml]) => parseRelationships(partPath, xml)),
    forecastSheetPresent: /<[^>]*sheet\b[^>]*name="Forecast"/i.test(workbookXml),
  };
}

function expectedComment(comment, parentId, fixture) {
  return {
    id: String(comment.id).toUpperCase(),
    parentId,
    personId: String(comment.personId).toUpperCase(),
    ref: fixture.address,
    date: normalizeIso(comment.date),
    done: Boolean(comment.done),
    text: comment.text,
  };
}

function sameComment(actual, expected) {
  return actual
    && actual.id === expected.id
    && actual.parentId === expected.parentId
    && actual.personId === expected.personId
    && actual.ref === expected.ref
    && actual.date === expected.date
    && actual.done === expected.done
    && actual.text === expected.text;
}

function samePerson(actual, expected) {
  return actual
    && actual.displayName === expected.author
    && actual.userId === expected.userId
    && actual.providerId === expected.providerId;
}

function hasRelationship(source, type, target, relationships) {
  return relationships.filter((relationship) => relationship.source === source
    && relationship.type === type
    && relationship.target === target
    && !relationship.external).length === 1;
}

export function nestedThreadedReplyGraphProfile(source) {
  const fixture = XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE;
  const expectedComments = [
    expectedComment(fixture.root, null, fixture),
    expectedComment(fixture.directReply, String(fixture.root.id).toUpperCase(), fixture),
    expectedComment(fixture.nestedReply, String(fixture.directReply.id).toUpperCase(), fixture),
  ];
  const commentsMatch = source.comments.length === expectedComments.length
    && expectedComments.every((expected, index) => sameComment(source.comments[index], expected));
  const people = new Map(source.people.map((person) => [person.id, person]));
  const peopleMatch = [fixture.root, fixture.directReply, fixture.nestedReply]
    .every((person) => samePerson(people.get(String(person.personId).toUpperCase()), person));
  const supportPartsMatch = JSON.stringify(source.threadedPaths) === JSON.stringify([fixture.threadedPartPath])
    && JSON.stringify(source.personPaths) === JSON.stringify([fixture.personPartPath])
    && source.contentTypes.get(fixture.threadedPartPath) === THREADED_COMMENTS_CONTENT_TYPE
    && source.contentTypes.get(fixture.personPartPath) === PERSON_CONTENT_TYPE;
  const relationshipsMatch = hasRelationship("xl/workbook.xml", PERSON_RELATIONSHIP_TYPE, fixture.personPartPath, source.relationships)
    && hasRelationship(fixture.worksheetPartPath, THREADED_COMMENTS_RELATIONSHIP_TYPE, fixture.threadedPartPath, source.relationships)
    && source.relationships.filter((relationship) => [PERSON_RELATIONSHIP_TYPE, THREADED_COMMENTS_RELATIONSHIP_TYPE].includes(relationship.type)).length === 2;
  return {
    ok: source.forecastSheetPresent
      && commentsMatch
      && peopleMatch
      && supportPartsMatch
      && relationshipsMatch
      && source.classicCommentPaths.length === 0,
    commentsMatch,
    peopleMatch,
    supportPartsMatch,
    relationshipsMatch,
    noClassicComments: source.classicCommentPaths.length === 0,
  };
}

function auditHash(audit, side) {
  const record = audit?.[side] || {};
  return String(record.sha256 || audit?.[`${side}Sha256`] || audit?.[`${side}_sha256`] || "");
}

function auditProvider(audit) {
  const provider = audit?.provider;
  if (typeof provider === "string") return provider;
  return [provider?.actual, provider?.selected, provider?.name, provider?.package, provider?.provider]
    .find((candidate) => typeof candidate === "string" && candidate.trim()) || "";
}

function auditFallbackIsFalse(audit) {
  const provider = audit?.provider || {};
  const values = [
    provider.silentFallback,
    provider.silent_fallback,
    provider.fallbackUsed,
    provider.fallback_used,
    provider.providerSwitched,
    provider.provider_switched,
    audit?.silentFallback,
    audit?.silent_fallback,
    audit?.fallbackUsed,
    audit?.fallback_used,
  ].filter((value) => value !== undefined);
  return values.length > 0 && values.every((value) => value === false || value === "false");
}

function someAuditField(value, predicate, depth = 0) {
  if (!value || typeof value !== "object" || depth > 6) return false;
  for (const [key, child] of Object.entries(value)) {
    if (predicate(key.replace(/[^a-z0-9]/gi, "").toLowerCase(), child)) return true;
    if (someAuditField(child, predicate, depth + 1)) return true;
  }
  return false;
}

function auditValueSucceeded(value) {
  return value === true || (value && typeof value === "object" && (
    value.ok === true
    || value.success === true
    || value.succeeded === true
    || value.completed === true
    || value.passed === true
  ));
}

function hasSuccessfulNamedAuditFact(audit, pattern) {
  return someAuditField(audit, (key, value) => pattern.test(key) && auditValueSucceeded(value));
}

function hasNamedZeroThreadProjection(audit) {
  return someAuditField(audit, (key, value) => value === 0
    && /(thread|comment|projection)/.test(key)
    && /(count|total|number)/.test(key));
}

function hasNamedNestedReplyBoundary(audit) {
  return someAuditField(audit, (key, value) => value === false && (
    key.includes("identitypreservingnestedreply")
    || key.includes("nestedorbranchedgraphseditable")
    || key.includes("targetthreadprojected")
  ));
}

function hasTypedNestedPreflight(audit) {
  const operations = Array.isArray(audit?.operations) ? audit.operations : [];
  const operationSucceeded = (pattern) => operations.some((operation) => pattern.test(String(operation?.name || operation?.operation || ""))
    && (auditValueSucceeded(operation) || /^(?:success|completed|passed)$/i.test(String(operation?.result || operation?.status || ""))));
  const recordedOperations = operations.some((operation) => /office.?kit/i.test(String(operation?.name || ""))
    && /import/i.test(String(operation?.name || ""))
    && /inspect/i.test(String(operation?.name || ""))
    && /^(?:completed|passed|unsupported_model_boundary|refused)$/i.test(String(operation?.result || ""))
    && operation?.mutation === false);
  const compoundOperation = operations.some((operation) => {
    const details = operation?.details || operation?.evidence || {};
    return /office.?kit/i.test(String(operation?.name || ""))
      && /import/i.test(String(operation?.name || ""))
      && /inspect/i.test(String(operation?.name || ""))
      && /^(?:completed|passed|unsupported_model_boundary|refused)$/i.test(String(operation?.result || ""))
      && (details?.imported_thread_count ?? details?.importedThreadCount) === 0
      && (details?.inspect_thread_item_count ?? details?.inspectThreadItemCount) === 0
      && (details?.workbook_verify_ok ?? details?.workbookVerifyOk) === true;
  });
  const preflight = audit?.preflight || {};
  const verification = audit?.verification || {};
  const structuredPreflight = preflight?.import?.ok === true
    && /(?:SpreadsheetFile\.)?importXlsx/i.test(String(preflight.import.method || ""))
    && preflight?.inspect?.ok === true
    && Number.isInteger(preflight.inspect.threadCount)
    && preflight.inspect.threadCount === 0
    && preflight.inspect.expectedNestedGraphProjected === false
    && preflight?.capabilityDecision?.supported === false;
  const validation = audit?.validation || {};
  const taskLocalTypedPreflight = validation?.import?.ok === true
    && validation?.inspect?.ok === true
    && /(?:^|,)\s*(?:workbook|sheet|thread)(?:\s*,|$)/i.test(String(validation.inspect.kind || ""))
    && validation?.targetThread?.ok === false
    && validation.targetThread?.expectedTarget?.sheetName === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && validation.targetThread?.expectedTarget?.address === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
    && validation.targetThread?.editableProjectionCount === 0
    && validation.targetThread?.modeledWorkbookThreadCount === 0
    && validation?.verify?.ok === true;
  const nestedRunnerPreflight = validation?.preflight;
  const runnerTypedPreflight = nestedRunnerPreflight?.ok === true
    && /(?:SpreadsheetFile\.)?importXlsx/i.test(String(nestedRunnerPreflight.method || ""))
    && nestedRunnerPreflight?.threadInspection?.target === `${XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName}!${XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address}`
    && nestedRunnerPreflight.threadInspection?.inspectRecordCount === 0
    && nestedRunnerPreflight.threadInspection?.publicModelMatchingThreadCount === 0
    && nestedRunnerPreflight.threadInspection?.identityPreservingNestedReplyCapability === false;
  const summaryTypedPreflight = validation?.import?.ok === true
    && validation?.inspect?.ok === true
    && /office[- ]?kit/i.test(String(validation.inspect.provider || auditProvider(audit)))
    && /"kind"\s*:\s*"workbook"/i.test(String(validation.inspect.summary || ""))
    && /"kind"\s*:\s*"sheet"/i.test(String(validation.inspect.summary || ""))
    && validation?.projectedThreadCount === 0
    && validation?.verify?.ok === true
    && validation?.sourceUnchanged?.ok === true;
  const operationTarget = audit?.operation?.target || {};
  const canonicalTarget = `${XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName}!${XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address}`;
  const operationTargetIsCanonical = operationTarget === canonicalTarget;
  const typedSourceBoundInspection = validation?.import?.ok === true
    && validation?.inspect?.ok === true
    && validation.inspect?.targetThreadCount === 0
    && validation.inspect?.editableProjectionAvailable === false
    && validation.inspect?.nestedReplyCount === 0
    && Array.isArray(validation.inspect?.comments)
    && validation.inspect.comments.length === 0
    && /"kind"\s*:\s*"sheet"[^\n]*"name"\s*:\s*"Forecast"/i.test(String(validation.inspect.sheetInspectionNdjson || ""))
    && operationTarget.sheet === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && operationTarget.cell === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
    && Array.isArray(audit?.operation?.performed)
    && audit.operation.performed.includes("office-kit-import")
    && audit.operation.performed.includes("office-kit-inspect");
  const agentPreflight = audit?.preflight;
  const agentPreflightWorkbook = Array.isArray(agentPreflight?.inspect?.workbook)
    ? agentPreflight.inspect.workbook
    : [];
  const auditedSourceBoundPreflight = agentPreflight?.import?.ok === true
    && /office[- ]?kit/i.test(String(agentPreflight.import.provider || auditProvider(audit)))
    && agentPreflight?.inspect?.ok === true
    && /(?:^|\.)workbook\.inspect$/i.test(String(agentPreflight.inspect.method || ""))
    && agentPreflightWorkbook.some((record) => record?.kind === "workbook")
    && agentPreflightWorkbook.some((record) => record?.kind === "sheet" && record?.name === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName)
    && Array.isArray(agentPreflight.inspect.target)
    && agentPreflight.inspect.target.length === 0
    && operationTarget.sheet === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && operationTarget.cell === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
    && audit?.operation?.attempted === false
    && audit?.operation?.refusal?.supported === false
    && validation?.sourceVerify?.ok === true
    && validation?.finalWorkbookArtifactAbsent === true;
  const targetBound = (
    operationTarget.sheet === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && operationTarget.cell === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
  ) || operationTargetIsCanonical || (
    validation?.targetThread?.expectedTarget?.sheetName === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && validation.targetThread?.expectedTarget?.address === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
  ) || (
    validation?.threadGraph?.target?.sheet === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && validation.threadGraph?.target?.cell === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
  ) || (
    validation?.inspect?.requestedTarget?.sheet === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && validation.inspect?.requestedTarget?.cell === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
  ) || (
    audit?.target?.sheet === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && audit.target?.cell === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
  ) || (
    audit?.target?.worksheet === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.sheetName
    && audit.target?.cell === XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.address
  ) || audit?.target?.address === canonicalTarget
    || audit?.target === canonicalTarget
    || nestedRunnerPreflight?.threadInspection?.target === canonicalTarget;
  const zeroProjectionCount = [
    validation?.inspect?.threadCount,
    validation?.inspect?.targetThreadCount,
    validation?.inspect?.threadCountAtTarget,
    validation?.inspect?.projectedThreadCount,
    validation?.inspect?.targetPublicThreadCount,
    validation?.projectedThreadCount,
    validation?.targetThread?.editableProjectionCount,
    validation?.targetThread?.modeledWorkbookThreadCount,
    validation?.targetThread?.count,
    validation?.threadGraph?.publicProjectionCountAtTarget,
    validation?.threadGraph?.totalPublicThreadCount,
    validation?.threadGraph?.threadCountAtTarget,
    nestedRunnerPreflight?.threadInspection?.inspectRecordCount,
    nestedRunnerPreflight?.threadInspection?.publicModelMatchingThreadCount,
    preflight?.inspect?.threadCount,
    preflight?.targetThread?.count,
    preflight?.targetThread?.commentCount,
    preflight?.targetPublicThreadCount,
  ].some((value) => value === 0);
  const zeroProjection = zeroProjectionCount || (
    validation?.targetThread?.found === false
    && validation.targetThread?.commentCount === 0
    && Array.isArray(validation.targetThread?.comments)
    && validation.targetThread.comments.length === 0
  ) || (
    preflight?.targetThread?.projectedByModel === false
    && preflight.targetThread?.count === 0
    && Array.isArray(preflight.targetThread?.comments)
    && preflight.targetThread.comments.length === 0
  ) || (
    Array.isArray(preflight?.inspect?.target)
    && preflight.inspect.target.length === 0
  ) || hasNamedZeroThreadProjection(audit);
  const nestedBoundaryCodes = [
    audit?.failure?.code,
    audit?.reason?.code,
    ...(Array.isArray(audit?.warnings) ? audit.warnings.map((warning) => warning?.code) : []),
  ].filter(Boolean).map((code) => String(code).toUpperCase());
  const explicitNestedBoundary = validation?.inspect?.expectedNestedGraphProjected === false
    || validation?.inspect?.editableProjectionAvailable === false
    || validation?.inspect?.sourceBoundThreadOmittedFromEditableModel === true
    || validation?.targetThread?.ok === false
    || validation?.threadGraph?.identityPreservingNestedMutationSupported === false
    || validation?.capability?.identityPreservingNestedReplyGraphMutationAvailable === false
    || (audit?.capability?.identityPreservingNestedReply === false && audit.capability?.sourceBound === true)
    || (validation?.capabilityBoundary?.supportedMutation === false && validation.capabilityBoundary?.sourceBound === true)
    || (validation?.inspect?.capabilityBoundary?.nestedOrBranchedGraphsEditable === false && validation.inspect.capabilityBoundary?.targetThreadProjected === false)
    || preflight?.targetThread?.projectedByModel === false
    || nestedBoundaryCodes.includes("SOURCE_BOUND_NESTED_THREADED_COMMENT_GRAPH")
    || nestedRunnerPreflight?.threadInspection?.identityPreservingNestedReplyCapability === false
    || hasNamedNestedReplyBoundary(audit);
  const inspectDocumented = validation?.inspect?.ok === true
    || validation?.inspect?.sourceBoundThreadOmittedFromEditableModel === true
    || preflight?.inspect?.ok === true
    || preflight?.inspectSucceeded === true
    || verification?.inspectSucceeded === true
    || (preflight?.workbookInspect?.ok === true && preflight?.targetInspect?.ok === true)
    || audit?.inspect?.ok === true
    || audit?.inspect?.succeeded === true
    || operationSucceeded(/inspect/i);
  const noMutationRecorded = audit?.operation?.performed !== true
    && audit?.operation?.attempted !== true
    && audit?.operation?.mutationAttempted !== true
    && audit?.operation?.attemptedMutation !== true
    && audit?.operation?.applied !== true
    && audit?.mutationAttempted !== true;
  const importDocumented = validation?.import?.ok === true
    || preflight?.import?.ok === true
    || preflight?.importSucceeded === true
    || verification?.importSucceeded === true
    || audit?.import?.ok === true
    || audit?.import?.succeeded === true
    || operationSucceeded(/(?:SpreadsheetFile\.)?importXlsx|\bimport\b/i)
    || hasSuccessfulNamedAuditFact(audit, /import/);
  const typedInspectDocumented = inspectDocumented
    || hasSuccessfulNamedAuditFact(audit, /inspect(?:ion)?/);
  const genericTypedPreflight = importDocumented
    && typedInspectDocumented
    && targetBound
    && zeroProjection
    && explicitNestedBoundary
    && noMutationRecorded;
  return /office[- ]?kit/i.test(auditProvider(audit))
    && (recordedOperations || compoundOperation || structuredPreflight || taskLocalTypedPreflight || runnerTypedPreflight || summaryTypedPreflight || typedSourceBoundInspection || auditedSourceBoundPreflight || genericTypedPreflight);
}

function usedTypedNestedInspection(commandText, audit) {
  const directImport = /(?:SpreadsheetFile\.)?importXlsx/i.test(commandText);
  // Task-local script names are intentionally unconstrained: the independent
  // typed audit, not a filename convention, proves that a public OfficeKit
  // preflight occurred. Mutation routes are rejected separately below.
  const tracedTask = /(?:^|\/)officekit(?:\.mjs)?\s+run\b|node_modules\/\.bin\/officekit\s+run\b|\bnode\s+(?:\.\/)?(?:[\w.-]+\/)*[\w.-]+\.mjs\b/i.test(commandText);
  return directImport || tracedTask && hasTypedNestedPreflight(audit);
}

function hasNestedThreadedMutationBypass(commandText) {
  return [
    /(?:SpreadsheetFile\.)?exportXlsx/i,
    /(?:SpreadsheetFile\.)?patchXlsx/i,
    /\b(?:JSZip|AdmZip|zipfile|openpyxl|xlsxwriter)\b/i,
    /(?:threadedcomments|persons)\/[^\n]{0,240}(?:write|replace|append|set|patch)/i,
    /\b(?:unzip|zip)\b/i,
  ].some((pattern) => pattern.test(commandText));
}

function failedClosedNestedThreadedAudit(audit, source) {
  if (!audit) return { present: false, valid: true };
  const output = audit.output;
  const noArtifact = output === null || output === undefined
    || output?.path === null || output?.path === undefined;
  const validation = audit?.validation || {};
  const noFallback = auditFallbackIsFalse(audit)
    || audit?.provider?.provider_switched === false
      && validation?.classic_note_fallback_used === false
      && validation?.flattened_root_reply_fallback_used === false
      && validation?.plain_cell_text_fallback_used === false;
  return {
    present: true,
    valid: audit?.status === "failed_closed"
      && auditHash(audit, "source") === source.sha256
      && noArtifact
      && audit?.deliveredModifiedXlsx !== true
      && audit?.delivered_modified_xlsx !== true
      && noFallback,
  };
}

export function gradeXlsxNestedThreadedReplyBoundaryEvidence({ evidence, audit, commands, item }) {
  const profile = nestedThreadedReplyGraphProfile(evidence.source);
  const commandText = commands.join("\n");
  const diagnostics = item.grade?.machine?.diagnosticTerms || [];
  const diagnosticText = `${evidence.finalMessage || ""}\n${audit?.reason || ""}\n${JSON.stringify(audit?.warnings || [])}`.toLowerCase();
  const auditResult = failedClosedNestedThreadedAudit(audit, evidence.source);
  const mutationBypass = hasNestedThreadedMutationBypass(commandText);
  return [
    check("xlsx-nested-machine:fixture-nested-reply-graph", "machine", profile.ok, {
      profile,
      comments: evidence.source.comments,
      people: evidence.source.people,
    }),
    check("xlsx-nested-machine:explicit-safe-refusal", "machine", diagnostics.some((term) => diagnosticText.includes(String(term).toLowerCase())), {
      expectedAny: diagnostics,
      actual: diagnosticText,
    }),
    check("xlsx-nested-machine:audit-is-failed-closed-when-present", "machine", auditResult.valid, {
      auditPresent: auditResult.present,
      audit: audit || null,
    }),
    gate("xlsx-nested-security:identity-topology-and-opc-graph-preserved", "security", profile.ok
      && evidence.source.comments[2]?.parentId === String(XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE.directReply.id).toUpperCase(), {
      profile,
    }),
    gate("xlsx-nested-security:no-note-downgrade-or-package-write", "security", !mutationBypass, {
      forbidden: "XLSX export/patch, direct ZIP access, or threaded-comments/person XML writes after a nested-reply refusal",
      commands,
    }),
    check("xlsx-nested-trace:typed-threaded-comment-inspection", "trace", usedTypedNestedInspection(commandText, audit), {
      expected: "public SpreadsheetFile.importXlsx or a traced task with a typed OfficeKit import/inspect audit",
      commands,
    }),
    gate("xlsx-nested-trace:no-silent-fallback", "trace", !mutationBypass, {
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

export async function gradeXlsxNestedThreadedReplyBoundaryCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const fixture = XLSX_THREADED_NESTED_REPLY_BOUNDARY_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  let source;
  try {
    source = await inspectXlsxNestedThreadedReplyGraph(path.join(workspace, "inputs", fixture.workbookName));
  } catch (error) {
    const checks = [
      gate("xlsx-nested-machine:readable-source", "machine", false, { error: error.message }),
      gate("xlsx-nested-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }
  const evidence = { source, finalMessage };
  const checks = gradeXlsxNestedThreadedReplyBoundaryEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}
