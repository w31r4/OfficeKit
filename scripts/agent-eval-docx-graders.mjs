import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  DOCX_CLASSIC_COMMENT_FIXTURE,
  DOCX_FOOTER_TEXT_FIXTURE,
  DOCX_HEADER_TEXT_FIXTURE,
  DOCX_MODERN_COMMENT_REPLY_BOUNDARY_FIXTURE,
} from "./agent-eval-office-fixtures.mjs";
import {
  gradeDocxSectionPageNumberingCase,
} from "./agent-eval-docx-section-page-numbering-grader.mjs";
export {
  gradeDocxSectionPageNumberingEvidence,
  inspectSectionPageNumberingDocx,
} from "./agent-eval-docx-section-page-numbering-grader.mjs";
import { renderOfficeFile } from "./agent-eval-office-native-render.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";

export const docxGradedCaseIds = new Set([
  "docx-classic-comment-text-edit",
  "docx-modern-comment-reply-boundary",
  "docx-header-text-edit",
  "docx-footer-text-edit",
  "docx-section-page-numbering-edit",
]);

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const SHIPPED_CLASSIC_COMMENT_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/documents|node_modules\/office-kit\/skills\/documents\/skills\/documents)\/examples\/officekit-classic-comment-edit-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_MODERN_COMMENT_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/documents|node_modules\/office-kit\/skills\/documents\/skills\/documents)\/examples\/officekit-modern-comment-thread-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_HEADER_TEXT_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/documents|node_modules\/office-kit\/skills\/documents\/skills\/documents)\/examples\/officekit-header-text-edit-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_FOOTER_TEXT_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/documents|node_modules\/office-kit\/skills\/documents\/skills\/documents)\/examples\/officekit-footer-text-edit-workflow\.mjs(?:$|[\s"'`])/i;

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

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeXmlText(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function tagIds(xml = "", name) {
  const pattern = new RegExp("<(?:[\\w.-]+:)?"+ name +"\\b[^>]*>", "gi");
  return [...String(xml).matchAll(pattern)]
    .map((match) => xmlAttributes(match[0]).id)
    .filter((id) => id !== undefined)
    .map(String);
}

function parseClassicComments(xml = "") {
  const comments = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?comment\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?comment>/g)) {
    const opening = /^<(?:[\w.-]+:)?comment\b[^>]*>/.exec(match[0])?.[0] || "";
    const attributes = xmlAttributes(opening);
    comments.push({
      id: String(attributes.id || ""),
      author: attributes.author || "",
      initials: attributes.initials || "",
      date: attributes.date || "",
      text: wordText(match[1]),
    });
  }
  return comments;
}

function parseParagraphs(xml = "") {
  const paragraphs = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?p>/g)) {
    const body = match[1];
    paragraphs.push({
      text: wordText(body),
      commentRangeStarts: tagIds(body, "commentRangeStart"),
      commentRangeEnds: tagIds(body, "commentRangeEnd"),
      commentReferences: tagIds(body, "commentReference"),
    });
  }
  return paragraphs;
}

export async function inspectClassicCommentDocx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const commentsPath = paths.find((name) => name.toLowerCase() === "word/comments.xml");
  const documentPath = paths.find((name) => name.toLowerCase() === "word/document.xml");
  const commentsXml = commentsPath ? await zip.file(commentsPath)?.async("text") || "" : "";
  const documentXml = documentPath ? await zip.file(documentPath)?.async("text") || "" : "";
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    commentsPath: commentsPath || null,
    documentPath: documentPath || null,
    commentsExtendedPaths: paths.filter((name) => /^word\/commentsExtended\.xml$/i.test(name)),
    peoplePaths: paths.filter((name) => /^word\/people\.xml$/i.test(name)),
    comments: parseClassicComments(commentsXml),
    paragraphs: parseParagraphs(documentXml),
  };
}

function parseModernCommentBodies(xml = "") {
  const comments = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?comment\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?comment>/g)) {
    const opening = /^<(?:[\w.-]+:)?comment\b[^>]*>/.exec(match[0])?.[0] || "";
    const paragraphs = [...match[1].matchAll(/<(?:[\w.-]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?p>/g)];
    const paragraphOpening = paragraphs.length === 1
      ? /^<(?:[\w.-]+:)?p\b[^>]*>/.exec(paragraphs[0][0])?.[0] || ""
      : "";
    const attributes = xmlAttributes(opening);
    const paragraphAttributes = xmlAttributes(paragraphOpening);
    comments.push({
      nativeId: String(attributes.id || ""),
      author: attributes.author || "",
      initials: attributes.initials || "",
      date: attributes.date || "",
      paraId: paragraphAttributes.paraId || "",
      paragraphCount: paragraphs.length,
      text: wordText(match[1]),
    });
  }
  return comments;
}

function parseOpenTags(xml = "", name) {
  const pattern = new RegExp(`<(?:[\\w.-]+:)?${escapeRegex(name)}\\b[^>]*\\/?\\s*>`, "gi");
  return [...String(xml).matchAll(pattern)].map((match) => xmlAttributes(match[0]));
}

function parseModernCommentPeople(xml = "") {
  const people = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?person\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?person>/g)) {
    const opening = /^<(?:[\w.-]+:)?person\b[^>]*>/.exec(match[0])?.[0] || "";
    const presence = parseOpenTags(match[1], "presenceInfo");
    const attributes = xmlAttributes(opening);
    people.push({
      author: attributes.author || "",
      providerId: presence.length === 1 ? presence[0].providerId || "" : "",
      userId: presence.length === 1 ? presence[0].userId || "" : "",
      presenceCount: presence.length,
    });
  }
  return people;
}

function xmlBoolean(value) {
  return /^(?:1|true)$/i.test(String(value || ""));
}

function normalizedRelationshipTarget(target = "") {
  const value = String(target || "").replaceAll("\\", "/");
  if (value.startsWith("/")) return value.slice(1);
  return path.posix.normalize(path.posix.join("word", value));
}

function relationshipRecords(xml = "") {
  return parseOpenTags(xml, "Relationship").map((attributes) => ({
    id: attributes.Id || "",
    type: attributes.Type || "",
    target: normalizedRelationshipTarget(attributes.Target),
  }));
}

function contentTypeOverrides(xml = "") {
  return parseOpenTags(xml, "Override").map((attributes) => ({
    partPath: String(attributes.PartName || "").replace(/^\//, ""),
    contentType: attributes.ContentType || "",
  }));
}

/**
 * Reads the modern Word-comment support graph directly from OPC/XML. It does
 * not call DocumentFile, so a candidate cannot satisfy the PromptBench oracle
 * simply by echoing its own model projection.
 */
export async function inspectModernCommentReplyGraphDocx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const entries = await Promise.all(paths.map(async (partPath) => [partPath, Buffer.from(await zip.file(partPath)?.async("uint8array") || [])]));
  const parts = new Map(entries.map(([partPath, value]) => [partPath, value.toString("utf8")]));
  const partHashes = Object.fromEntries(entries.map(([partPath, value]) => [partPath, sha256(value)]));
  const text = (partPath) => parts.get(partPath) || "";
  const comments = parseModernCommentBodies(text("word/comments.xml"));
  const commentsExtended = parseOpenTags(text("word/commentsExtended.xml"), "commentEx").map((attributes) => ({
    paraId: attributes.paraId || "",
    parentParaId: attributes.paraIdParent || "",
    resolved: xmlBoolean(attributes.done),
  }));
  const commentsIds = parseOpenTags(text("word/commentsIds.xml"), "commentId").map((attributes) => ({
    paraId: attributes.paraId || "",
    durableId: attributes.durableId || "",
  }));
  const commentsExtensible = parseOpenTags(text("word/commentsExtensible.xml"), "commentExtensible").map((attributes) => ({
    durableId: attributes.durableId || "",
    dateUtc: attributes.dateUtc || "",
  }));
  const documentXml = text("word/document.xml");
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes,
    comments,
    commentsExtended,
    commentsIds,
    commentsExtensible,
    people: parseModernCommentPeople(text("word/people.xml")),
    paragraphs: parseParagraphs(documentXml),
    relationships: relationshipRecords(text("word/_rels/document.xml.rels")),
    contentTypeOverrides: contentTypeOverrides(text("[Content_Types].xml")),
  };
}

export function modernCommentReplyGraphProfile(document) {
  const fixture = DOCX_MODERN_COMMENT_REPLY_BOUNDARY_FIXTURE;
  const expectedParts = [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
    "word/people.xml",
  ];
  const byNativeId = new Map(fixture.comments.map((comment) => [comment.nativeId, comment]));
  const commentsMatch = document.comments.length === fixture.comments.length
    && document.comments.every((actual, index) => {
      const expected = fixture.comments[index];
      return actual.nativeId === expected.nativeId
        && actual.paraId === expected.paraId
        && actual.author === expected.author
        && actual.initials === expected.initials
        && actual.date === expected.date
        && actual.text === expected.text
        && actual.paragraphCount === 1;
    });
  const extendedMatch = document.commentsExtended.length === fixture.comments.length
    && document.commentsExtended.every((actual, index) => {
      const expected = fixture.comments[index];
      const parent = expected.parentNativeId ? byNativeId.get(expected.parentNativeId) : null;
      return actual.paraId === expected.paraId
        && actual.parentParaId === (parent?.paraId || "")
        && actual.resolved === expected.resolved;
    });
  const idsMatch = document.commentsIds.length === fixture.comments.length
    && document.commentsIds.every((actual, index) => actual.paraId === fixture.comments[index].paraId
      && actual.durableId === fixture.comments[index].durableId);
  const extensibleMatch = document.commentsExtensible.length === fixture.comments.length
    && document.commentsExtensible.every((actual, index) => actual.durableId === fixture.comments[index].durableId
      && actual.dateUtc === fixture.comments[index].dateUtc);
  const peopleMatch = document.people.length === fixture.comments.length
    && document.people.every((actual, index) => actual.author === fixture.comments[index].author
      && actual.providerId === fixture.comments[index].providerId
      && actual.userId === fixture.comments[index].userId
      && actual.presenceCount === 1);
  const anchors = document.paragraphs.filter((paragraph) => paragraph.text === fixture.anchorText);
  const root = fixture.comments[0];
  const anchorMatch = anchors.length === 1
    && anchors[0].commentRangeStarts.filter((id) => id === root.nativeId).length === 1
    && anchors[0].commentRangeEnds.filter((id) => id === root.nativeId).length === 1
    && anchors[0].commentReferences.filter((id) => id === root.nativeId).length === 1
    && !document.paragraphs.some((paragraph) => ["1", "2"].some((id) => paragraph.commentRangeStarts.includes(id)
      || paragraph.commentRangeEnds.includes(id) || paragraph.commentReferences.includes(id)));
  const expectedSupportParts = new Map([
    ["word/comments.xml", {
      relationshipType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
    }],
    ["word/commentsExtended.xml", {
      relationshipType: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
    }],
    ["word/commentsIds.xml", {
      relationshipType: "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
    }],
    ["word/commentsExtensible.xml", {
      relationshipType: "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
    }],
    ["word/people.xml", {
      relationshipType: "http://schemas.microsoft.com/office/2011/relationships/people",
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",
    }],
  ]);
  const supportPartsMatch = expectedParts.every((partPath) => {
    const expected = expectedSupportParts.get(partPath);
    return document.paths.includes(partPath)
      && document.contentTypeOverrides.some((override) => override.partPath === partPath && override.contentType === expected.contentType)
      && document.relationships.some((relationship) => relationship.target === partPath && relationship.type === expected.relationshipType);
  });
  const bodyMatch = document.paragraphs.some((paragraph) => paragraph.text === fixture.title)
    && document.paragraphs.some((paragraph) => paragraph.text === fixture.supportingText);
  return {
    ok: commentsMatch && extendedMatch && idsMatch && extensibleMatch && peopleMatch && anchorMatch && supportPartsMatch && bodyMatch,
    commentsMatch,
    extendedMatch,
    idsMatch,
    extensibleMatch,
    peopleMatch,
    anchorMatch,
    supportPartsMatch,
    bodyMatch,
  };
}

function normalizeTargetText(xml, expectedText) {
  const pattern = new RegExp(`(<w:t(?:\\s[^>]*)?>)${escapeRegex(escapeXmlText(expectedText))}(</w:t>)`, "g");
  let matches = 0;
  const normalized = xml.replace(pattern, (_whole, open, close) => {
    matches += 1;
    return `${open}__OFFICE_KIT_TARGET_TEXT__${close}`;
  });
  return { normalized, matches };
}

function headerFooterPartRecord(partPath, xml) {
  return {
    path: partPath,
    xml,
    paragraphs: parseParagraphs(xml).map((paragraph) => paragraph.text),
  };
}

/**
 * Independent package-level reader for the source-bound header-edit case.
 * It intentionally does not use DocumentFile so a wrapper cannot satisfy the
 * black-box evaluator by merely repeating its own model projection.
 */
export async function inspectPageFurnitureTextDocx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const entries = await Promise.all(paths.map(async (partPath) => [partPath, Buffer.from(await zip.file(partPath)?.async("uint8array") || [])]));
  const partHashes = Object.fromEntries(entries.map(([partPath, value]) => [partPath, sha256(value)]));
  const content = new Map(entries.map(([partPath, value]) => [partPath, value.toString("utf8")]));
  const documentPath = paths.find((name) => name.toLowerCase() === "word/document.xml") || null;
  const headerParts = paths.filter((name) => /^word\/header\d+\.xml$/i.test(name))
    .map((partPath) => headerFooterPartRecord(partPath, content.get(partPath) || ""));
  const footerParts = paths.filter((name) => /^word\/footer\d+\.xml$/i.test(name))
    .map((partPath) => headerFooterPartRecord(partPath, content.get(partPath) || ""));
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes,
    documentPath,
    bodyParagraphs: parseParagraphs(documentPath ? content.get(documentPath) || "" : "").map((paragraph) => paragraph.text),
    headerParts,
    footerParts,
  };
}

export const inspectHeaderTextDocx = inspectPageFurnitureTextDocx;
export const inspectFooterTextDocx = inspectPageFurnitureTextDocx;

function auditProvider(audit) {
  const provider = audit?.provider;
  if (typeof provider === "string") return provider;
  for (const value of [provider?.actual, provider?.selected, provider?.name, provider?.package, provider?.provider]) {
    if (typeof value === "string" && value.trim()) return value;
  }
  return "";
}

function auditVersion(audit) {
  const provider = audit?.provider;
  return String(provider?.version || audit?.providerVersion || "");
}

function auditFallbackIsFalse(audit) {
  const provider = audit?.provider || {};
  const values = [
    provider.silentFallback,
    provider.silent_fallback,
    provider.fallbackUsed,
    provider.fallback_used,
    audit?.silentFallback,
    audit?.silent_fallback,
    audit?.fallbackUsed,
    audit?.fallback_used,
  ].filter((value) => value !== undefined);
  return values.length > 0 && values.every((value) => value === false || value === "false");
}

function auditStrategy(audit) {
  const policy = audit?.savePolicy || audit?.save_policy || audit?.saveStrategy || audit?.save_strategy;
  return String(typeof policy === "string" ? policy : policy?.strategy || policy?.selected || audit?.strategy || "");
}

function auditOperation(audit) {
  const operation = audit?.operation;
  return String(typeof operation === "string" ? operation : operation?.type || operation?.name || operation?.operation || "");
}

function auditHash(audit, side) {
  const record = audit?.[side] || {};
  return String(record.sha256 || audit?.[side + "Sha256"] || audit?.[side + "_sha256"] || "");
}

function sameCommentMetadata(left, right) {
  return left && right
    && left.id === right.id
    && left.author === right.author
    && left.initials === right.initials
    && left.date === right.date;
}

function anchoredParagraph(document, commentId, anchorText) {
  const matches = document.paragraphs.filter((paragraph) => paragraph.text === anchorText);
  if (matches.length !== 1) return null;
  const paragraph = matches[0];
  const id = String(commentId);
  return {
    paragraph,
    allMarkersPresent: paragraph.commentRangeStarts.includes(id)
      && paragraph.commentRangeEnds.includes(id)
      && paragraph.commentReferences.includes(id),
  };
}

function paragraphsPreserved(source, output) {
  return JSON.stringify(source.paragraphs.map((paragraph) => paragraph.text))
    === JSON.stringify(output.paragraphs.map((paragraph) => paragraph.text));
}

function visualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount;
  const pixelsStable = pageCountsMatch
    && source?.pages?.length === output?.pages?.length
    && source?.pages?.every((page, index) => {
      const next = output.pages[index];
      return next
        && page.width === next.width
        && page.height === next.height
        && page.pixelSha256 === next.pixelSha256;
    });
  return { available, rendered, pageCountsMatch, pixelsStable };
}

function usedTypedDocxRoundTrip(commandText) {
  const directPublicApi = /(?:DocumentFile\.)?importDocx/i.test(commandText)
    && /(?:DocumentFile\.)?exportDocx/i.test(commandText);
  return directPublicApi || SHIPPED_CLASSIC_COMMENT_WORKFLOW.test(commandText);
}

export function gradeDocxClassicCommentEvidence({ evidence, audit, commands }) {
  const fixture = DOCX_CLASSIC_COMMENT_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const sourceComment = source.comments[0];
  const outputComment = output.comments[0];
  const sourceAnchor = anchoredParagraph(source, sourceComment?.id, fixture.anchorText);
  const outputAnchor = anchoredParagraph(output, outputComment?.id, fixture.anchorText);
  const visual = visualEvidence(evidence.visual?.source, evidence.visual?.output);
  const commandText = commands.join("\n");
  const classicTopology = source.commentsPath === "word/comments.xml"
    && output.commentsPath === "word/comments.xml"
    && source.comments.length === 1
    && output.comments.length === 1
    && source.commentsExtendedPaths.length === 0
    && output.commentsExtendedPaths.length === 0
    && source.peoplePaths.length === 0
    && output.peoplePaths.length === 0;
  return [
    check("docx-machine:fixture-classic-comment", "machine", Boolean(sourceComment)
      && sourceComment.author === fixture.comment.author
      && sourceComment.initials === fixture.comment.initials
      && sourceComment.date === fixture.comment.date
      && sourceComment.text === fixture.comment.originalText, { sourceComment }),
    check("docx-machine:comment-text-edited", "machine", sameCommentMetadata(sourceComment, outputComment)
      && outputComment?.text === fixture.comment.replacementText, { sourceComment, outputComment }),
    check("docx-machine:comment-anchor-preserved", "machine", sourceAnchor?.allMarkersPresent === true
      && outputAnchor?.allMarkersPresent === true
      && sourceComment?.id === outputComment?.id, {
      sourceAnchor,
      outputAnchor,
      sourceCommentId: sourceComment?.id,
      outputCommentId: outputComment?.id,
    }),
    check("docx-machine:visible-document-text-preserved", "machine", paragraphsPreserved(source, output), {
      source: source.paragraphs.map((paragraph) => paragraph.text),
      output: output.paragraphs.map((paragraph) => paragraph.text),
    }),
    check("docx-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("docx-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("docx-visual:comment-only-body-stable", "visual", visual.pixelsStable, {
      visual: evidence.visual,
      note: "Headless LibreOffice does not render review balloons reliably; classic-comment structure is checked independently.",
    }),
    gate("docx-security:classic-topology-and-modern-graph-absence", "security", classicTopology, {
      commentsPath: { source: source.commentsPath, output: output.commentsPath },
      commentsExtendedPaths: { source: source.commentsExtendedPaths, output: output.commentsExtendedPaths },
      peoplePaths: { source: source.peoplePaths, output: output.peoplePaths },
      counts: { source: source.comments.length, output: output.comments.length },
    }),
    gate("docx-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("docx-trace:office-kit-provider", "trace", /office[- ]?kit/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("docx-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("docx-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("docx-trace:classic-comment-operation", "trace", /classic|comment/i.test(auditOperation(audit)), {
      operation: auditOperation(audit),
    }),
    check("docx-trace:typed-roundtrip", "trace", usedTypedDocxRoundTrip(commandText), {
      expected: "public DocumentFile importDocx/exportDocx calls or the integrity-protected published classic-comment workflow",
    }),
    check("docx-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

function modernCommentAuditHasTypedPreflight(audit) {
  const operations = Array.isArray(audit?.operations) ? audit.operations : [];
  const recordedOperations = operations.some((operation) => operation?.name === "officekit_import" && operation?.result === "completed" && operation?.mutation === false)
    && operations.some((operation) => operation?.name === "officekit_inspect" && operation?.result === "completed" && operation?.mutation === false);
  const compoundOperation = operations.some((operation) => {
    const name = String(operation?.name || "");
    const details = operation?.details;
    return /office[- ]?kit/i.test(name)
      && /import/i.test(name)
      && /inspect/i.test(name)
      && ["completed", "unsupported_model_boundary", "refused"].includes(String(operation?.result || ""))
      && (details?.imported_comment_count ?? details?.importedCommentCount) === 0
      && (details?.inspect_comment_item_count ?? details?.inspectCommentItemCount) === 0
      && (details?.document_verify_ok ?? details?.documentVerifyOk) === true;
  });
  const preflight = audit?.preflight;
  const structuredPreflight = preflight?.import?.ok === true
    && /(?:DocumentFile\.)?importDocx/i.test(String(preflight.import.method || ""))
    && preflight?.inspect?.ok === true
    && Number.isInteger(preflight.inspect.commentCount)
    && preflight.inspect.commentCount === 0
    && preflight.inspect.expectedNestedGraphProjected === false
    && preflight?.capabilityDecision?.supported === false;
  return /office[- ]?kit/i.test(auditProvider(audit)) && (recordedOperations || compoundOperation || structuredPreflight);
}

function usedTypedModernCommentInspection(commandText, audit) {
  return /(?:DocumentFile\.)?importDocx/i.test(commandText)
    || SHIPPED_MODERN_COMMENT_WORKFLOW.test(commandText)
    || /(?:^|\/)officekit(?:\.mjs)?\s+run\b|node_modules\/\.bin\/officekit\s+run\b/i.test(commandText) && modernCommentAuditHasTypedPreflight(audit);
}

function hasModernCommentMutationBypass(commandText) {
  return [
    /(?:DocumentFile\.)?exportDocx/i,
    /(?:DocumentFile\.)?patchDocx/i,
    /\b(?:JSZip|AdmZip|zipfile)\b/i,
    /(?:commentsExtended|commentsExtensible|commentsIds|people)\.xml[^\n]{0,240}(?:write|replace|append|set|patch)/i,
    /\bzip\s+-[a-z]*[udf][a-z]*\b/i,
  ].some((pattern) => pattern.test(commandText));
}

function failedClosedModernCommentAudit(audit, source) {
  if (!audit) return { present: false, valid: true };
  const output = audit.output;
  const noArtifact = output === null || output === undefined
    || output?.path === null || output?.path === undefined;
  const noFallback = auditFallbackIsFalse(audit)
    || audit?.provider?.provider_switch === false
      && audit?.validation?.classic_comment_fallback_used === false
      && audit?.validation?.plain_text_fallback_used === false
      && audit?.validation?.new_root_fallback_used === false;
  return {
    present: true,
    valid: audit?.status === "failed_closed"
      && auditHash(audit, "source") === source.sha256
      && noArtifact
      && audit?.deliveredModifiedDocx !== true
      && audit?.delivered_modified_docx !== true
      && noFallback,
  };
}

export function gradeDocxModernCommentReplyBoundaryEvidence({ evidence, audit, commands, item }) {
  const profile = modernCommentReplyGraphProfile(evidence.source);
  const commandText = commands.join("\n");
  const diagnostics = item.grade?.machine?.diagnosticTerms || [];
  const diagnosticText = `${evidence.finalMessage || ""}\n${audit?.reason || ""}\n${JSON.stringify(audit?.warnings || [])}`.toLowerCase();
  const auditResult = failedClosedModernCommentAudit(audit, evidence.source);
  const mutationBypass = hasModernCommentMutationBypass(commandText);
  return [
    check("docx-modern-machine:fixture-nested-reply-graph", "machine", profile.ok, {
      profile,
      comments: evidence.source.comments,
      commentsExtended: evidence.source.commentsExtended,
    }),
    check("docx-modern-machine:explicit-safe-refusal", "machine", diagnostics.some((term) => diagnosticText.includes(String(term).toLowerCase())), {
      expectedAny: diagnostics,
      actual: diagnosticText,
    }),
    check("docx-modern-machine:audit-is-failed-closed-when-present", "machine", auditResult.valid, {
      auditPresent: auditResult.present,
      audit: audit || null,
    }),
    gate("docx-modern-security:modern-identity-and-reply-of-reply-preserved", "security", profile.ok
      && evidence.source.commentsExtended[2]?.parentParaId === DOCX_MODERN_COMMENT_REPLY_BOUNDARY_FIXTURE.comments[1].paraId, {
      profile,
    }),
    gate("docx-modern-security:no-classic-downgrade-or-package-write", "security", !mutationBypass, {
      forbidden: "DOCX export/patch or low-level comments XML write after a nested-reply refusal",
      commands,
    }),
    check("docx-modern-trace:typed-modern-comment-inspection", "trace", usedTypedModernCommentInspection(commandText, audit), {
      expected: "public DocumentFile.importDocx, the published modern-comment workflow, or a traced officekit run with a typed audit preflight",
      commands,
    }),
    gate("docx-modern-trace:no-silent-fallback", "trace", !mutationBypass, {
      commands,
    }),
  ];
}

function pageFurniturePartWithText(document, kind, text) {
  const parts = kind === "header" ? document.headerParts : document.footerParts;
  const matches = parts.filter((part) => part.paragraphs.filter((paragraph) => paragraph === text).length === 1);
  return matches.length === 1 ? matches[0] : null;
}

function changedPackageParts(source, output) {
  if (JSON.stringify(source.paths) !== JSON.stringify(output.paths)) return null;
  return source.paths.filter((partPath) => source.partHashes[partPath] !== output.partHashes[partPath]);
}

function normalizedPageFurnitureResidual(sourcePart, outputPart, sourceText, replacementText) {
  if (!sourcePart || !outputPart || sourcePart.path !== outputPart.path) return { ok: false, source: null, output: null };
  const source = normalizeTargetText(sourcePart.xml, sourceText);
  const output = normalizeTargetText(outputPart.xml, replacementText);
  return {
    ok: source.matches === 1 && output.matches === 1 && source.normalized === output.normalized,
    source,
    output,
  };
}

function pageFurnitureVisibleChange(source, output) {
  const visual = visualEvidence(source, output);
  const pixelsChanged = visual.pageCountsMatch
    && source?.pages?.length === output?.pages?.length
    && source.pages.some((page, index) => page.pixelSha256 !== output.pages[index]?.pixelSha256);
  return { ...visual, pixelsChanged };
}

function usedTypedPageFurnitureRoundTrip(commandText, kind) {
  const directPublicApi = /(?:DocumentFile\.)?importDocx/i.test(commandText)
    && /(?:DocumentFile\.)?exportDocx/i.test(commandText);
  const workflow = kind === "header" ? SHIPPED_HEADER_TEXT_WORKFLOW : SHIPPED_FOOTER_TEXT_WORKFLOW;
  return directPublicApi || workflow.test(commandText);
}

function gradeDocxPageFurnitureTextEvidence({ evidence, audit, commands, fixture, kind }) {
  const profile = fixture[kind];
  const otherKind = kind === "header" ? "footer" : "header";
  const targetPartsKey = kind === "header" ? "headerParts" : "footerParts";
  const otherPartsKey = otherKind === "header" ? "headerParts" : "footerParts";
  const prefix = `docx-${kind}`;
  const source = evidence.source;
  const output = evidence.output;
  const sourceTarget = pageFurniturePartWithText(source, kind, profile.originalText);
  const outputTarget = pageFurniturePartWithText(output, kind, profile.replacementText);
  const residual = normalizedPageFurnitureResidual(sourceTarget, outputTarget, profile.originalText, profile.replacementText);
  const changed = changedPackageParts(source, output);
  const companionStable = Boolean(sourceTarget && outputTarget)
    && JSON.stringify(sourceTarget.paragraphs.map((paragraph) => paragraph === profile.originalText ? "__target__" : paragraph))
      === JSON.stringify(outputTarget.paragraphs.map((paragraph) => paragraph === profile.replacementText ? "__target__" : paragraph));
  const otherStable = source[otherPartsKey].length === output[otherPartsKey].length
    && source[otherPartsKey].every((part, index) => part.path === output[otherPartsKey][index]?.path
      && source.partHashes[part.path] === output.partHashes[part.path]);
  const visual = pageFurnitureVisibleChange(evidence.visual?.source, evidence.visual?.output);
  const commandText = commands.join("\n");
  const auditTarget = audit?.operation?.target || {};
  return [
    check(`${prefix}-machine:fixture-source-profile`, "machine", Boolean(sourceTarget)
      && sourceTarget.paragraphs.includes(profile.companionText)
      && source[otherPartsKey].length === 1
      && source.bodyParagraphs.includes(fixture.title)
      && fixture.body.every((paragraph) => source.bodyParagraphs.includes(paragraph)), {
      sourceTarget,
      [otherPartsKey]: source[otherPartsKey].map((part) => ({ path: part.path, paragraphs: part.paragraphs })),
      bodyParagraphs: source.bodyParagraphs,
    }),
    check(`${prefix}-machine:requested-text-edited`, "machine", Boolean(sourceTarget && outputTarget)
      && outputTarget.paragraphs.includes(profile.replacementText)
      && !outputTarget.paragraphs.includes(profile.originalText)
      && companionStable, {
      sourceTarget,
      outputTarget,
      companionStable,
    }),
    check(`${prefix}-machine:body-and-${otherKind}-stable`, "machine", JSON.stringify(source.bodyParagraphs) === JSON.stringify(output.bodyParagraphs)
      && otherStable, {
      body: { source: source.bodyParagraphs, output: output.bodyParagraphs },
      [`${otherKind}Stable`]: otherStable,
    }),
    check(`${prefix}-machine:${kind}-residual-stable`, "machine", residual.ok, {
      sourceMatches: residual.source?.matches ?? 0,
      outputMatches: residual.output?.matches ?? 0,
    }),
    check(`${prefix}-machine:audit-succeeded`, "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check(`${prefix}-visual:native-render`, "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check(`${prefix}-visual:${kind}-change-visible`, "visual", visual.pixelsChanged, {
      visual: evidence.visual,
      note: "The package gate proves the changed part; native rendering confirms that the requested page-furniture edit remains visible.",
    }),
    gate(`${prefix}-security:only-target-${kind}-part-changed`, "security", Boolean(sourceTarget)
      && JSON.stringify(changed) === JSON.stringify([sourceTarget.path]), {
      changed,
      targetPart: sourceTarget?.path || null,
    }),
    gate(`${prefix}-security:${otherKind}-field-and-package-inventory-preserved`, "security", otherStable
      && changed !== null
      && source[otherPartsKey].some((part) => part.paragraphs.includes(fixture[otherKind].text)), {
      [`${otherKind}Stable`]: otherStable,
      sourcePaths: source.paths,
      outputPaths: output.paths,
    }),
    gate(`${prefix}-security:byte-bound-audit-provenance`, "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check(`${prefix}-trace:office-kit-provider`, "trace", /office[- ]?kit/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate(`${prefix}-trace:no-silent-fallback`, "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check(`${prefix}-trace:rewrite-policy`, "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check(`${prefix}-trace:source-bound-${kind}-operation`, "trace", new RegExp(kind, "i").test(auditOperation(audit))
      && auditTarget.sectionIndex === profile.sectionIndex
      && auditTarget.referenceType === profile.referenceType
      && auditTarget.partPath === sourceTarget?.path, {
      operation: audit?.operation || null,
      expected: { sectionIndex: profile.sectionIndex, referenceType: profile.referenceType, partPath: sourceTarget?.path || null },
    }),
    check(`${prefix}-trace:typed-roundtrip`, "trace", usedTypedPageFurnitureRoundTrip(commandText, kind), {
      expected: `public DocumentFile importDocx/exportDocx calls or the integrity-protected published ${kind}-text workflow`,
    }),
    check(`${prefix}-trace:second-import`, "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

export function gradeDocxHeaderTextEvidence(options) {
  return gradeDocxPageFurnitureTextEvidence({ ...options, fixture: DOCX_HEADER_TEXT_FIXTURE, kind: "header" });
}

export function gradeDocxFooterTextEvidence(options) {
  return gradeDocxPageFurnitureTextEvidence({ ...options, fixture: DOCX_FOOTER_TEXT_FIXTURE, kind: "footer" });
}

async function readAudit(workspace) {
  try {
    return JSON.parse(await fs.readFile(path.join(workspace, "outputs", "audit.json"), "utf8"));
  } catch {
    return null;
  }
}

async function gradeDocxClassicCommentCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const fixture = DOCX_CLASSIC_COMMENT_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const sourcePath = path.join(workspace, "inputs", fixture.documentName);
  const outputPath = path.join(workspace, "outputs", "legal-review-updated.docx");
  let source;
  let output;
  try {
    [source, output] = await Promise.all([
      inspectClassicCommentDocx(sourcePath),
      inspectClassicCommentDocx(outputPath),
    ]);
  } catch (error) {
    const checks = [
      gate("docx-machine:readable-output", "machine", false, { error: error.message }),
      gate("docx-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }

  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(sourcePath, "docx-source"),
    renderOfficeFile(outputPath, "docx-output"),
  ]);
  const visualUnavailable = [sourceRender, outputRender].find((result) => !result.available);
  if (visualUnavailable) {
    return {
      supported: true,
      graded: false,
      checks: [],
      evidence: { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage },
      pending: ["native LibreOffice/Poppler document rendering"],
      infrastructureErrors: [visualUnavailable.reason],
    };
  }

  const evidence = { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage };
  const checks = gradeDocxClassicCommentEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

async function gradeDocxModernCommentReplyBoundaryCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const fixture = DOCX_MODERN_COMMENT_REPLY_BOUNDARY_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const sourcePath = path.join(workspace, "inputs", fixture.documentName);
  let source;
  try {
    source = await inspectModernCommentReplyGraphDocx(sourcePath);
  } catch (error) {
    const checks = [
      gate("docx-modern-machine:readable-source", "machine", false, { error: error.message }),
      gate("docx-modern-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }
  // No modified DOCX is allowed for this refusal case. Native page rendering is
  // not a semantic oracle for Word comments, and there is no output package to
  // compare, so structural support-part and source-immutability checks lead.
  const evidence = { source, finalMessage };
  const checks = gradeDocxModernCommentReplyBoundaryEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

async function gradeDocxPageFurnitureTextCase({
  item,
  workspace,
  finalMessage,
  trace,
  weights = defaultWeights,
  fixture,
  kind,
  outputName,
}) {
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const sourcePath = path.join(workspace, "inputs", fixture.documentName);
  const outputPath = path.join(workspace, "outputs", outputName);
  const prefix = `docx-${kind}`;
  let source;
  let output;
  try {
    [source, output] = await Promise.all([
      inspectPageFurnitureTextDocx(sourcePath),
      inspectPageFurnitureTextDocx(outputPath),
    ]);
  } catch (error) {
    const checks = [
      gate(`${prefix}-machine:readable-output`, "machine", false, { error: error.message }),
      gate(`${prefix}-security:no-partial-success`, "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }

  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(sourcePath, `docx-${kind}-source`),
    renderOfficeFile(outputPath, `docx-${kind}-output`),
  ]);
  const visualUnavailable = [sourceRender, outputRender].find((result) => !result.available);
  if (visualUnavailable) {
    return {
      supported: true,
      graded: false,
      checks: [],
      evidence: { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage },
      pending: ["native LibreOffice/Poppler document rendering"],
      infrastructureErrors: [visualUnavailable.reason],
    };
  }

  const evidence = { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage };
  const checks = gradeDocxPageFurnitureTextEvidence({ evidence, audit, commands, item, fixture, kind });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

async function gradeDocxHeaderTextCase(options) {
  return gradeDocxPageFurnitureTextCase({
    ...options,
    fixture: DOCX_HEADER_TEXT_FIXTURE,
    kind: "header",
    outputName: "board-brief-header-reviewed.docx",
  });
}

async function gradeDocxFooterTextCase(options) {
  return gradeDocxPageFurnitureTextCase({
    ...options,
    fixture: DOCX_FOOTER_TEXT_FIXTURE,
    kind: "footer",
    outputName: "board-brief-footer-reviewed.docx",
  });
}

export async function gradeDocxCase(options) {
  if (options.item.id === "docx-classic-comment-text-edit") return gradeDocxClassicCommentCase(options);
  if (options.item.id === "docx-modern-comment-reply-boundary") return gradeDocxModernCommentReplyBoundaryCase(options);
  if (options.item.id === "docx-header-text-edit") return gradeDocxHeaderTextCase(options);
  if (options.item.id === "docx-footer-text-edit") return gradeDocxFooterTextCase(options);
  if (options.item.id === "docx-section-page-numbering-edit") return gradeDocxSectionPageNumberingCase(options);
  return { supported: false };
}
