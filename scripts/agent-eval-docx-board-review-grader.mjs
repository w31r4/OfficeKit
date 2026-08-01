import crypto from "node:crypto";
import fs from "node:fs/promises";

import JSZip from "jszip";

import { DOCX_SURGICAL_BOARD_REVIEW_FIXTURE } from "./agent-eval-office-fixtures.mjs";

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function decodeXml(value = "") {
  return String(value)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
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

function parseComments(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?comment\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?comment>/g)].map((match) => ({
    ...xmlAttributes(/^<(?:[\w.-]+:)?comment\b[^>]*>/.exec(match[0])?.[0] || ""),
    text: wordText(match[1]),
  }));
}

function parseParagraphs(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?p>/g)].map((match) => ({
    text: wordText(match[1]),
    commentIds: [...match[1].matchAll(/<(?:[\w.-]+:)?comment(?:RangeStart|RangeEnd|Reference)\b[^>]*>/g)]
      .map((entry) => xmlAttributes(entry[0]).id)
      .filter(Boolean),
  }));
}

function topLevelXmlElements(xml = "", localName) {
  const token = new RegExp(`<(?:(?:[\\w.-]+):)?${localName}\\b[^>]*>|</(?:(?:[\\w.-]+):)?${localName}\\s*>`, "gi");
  const elements = [];
  let depth = 0;
  let start = -1;
  for (const match of String(xml).matchAll(token)) {
    if (/^<\//.test(match[0])) {
      depth -= 1;
      if (depth === 0 && start >= 0) {
        elements.push(String(xml).slice(start, (match.index || 0) + match[0].length));
        start = -1;
      }
    } else {
      if (depth === 0) start = match.index || 0;
      depth += 1;
    }
  }
  return elements;
}

function tableRows(table = "") {
  return topLevelXmlElements(table, "tr").map((row) => topLevelXmlElements(row, "tc").map((cell) => wordText(cell)));
}

function changedParts(source, output) {
  const all = [...new Set([...Object.keys(source.partHashes), ...Object.keys(output.partHashes)])].sort();
  return all.filter((part) => source.partHashes[part] !== output.partHashes[part]);
}

function check(id, category, passed, details = {}) {
  return { id, category, gate: false, passed: Boolean(passed), ...details };
}

function gate(id, category, passed, details = {}) {
  return { id, category, gate: true, passed: Boolean(passed), ...details };
}

function auditProvider(audit) {
  const provider = audit?.provider;
  return String(typeof provider === "string" ? provider : provider?.actual || provider?.selected || provider?.name || "");
}

function auditVersion(audit) {
  return String(audit?.provider?.version || audit?.providerVersion || "");
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

function auditFallbackIsFalse(audit) {
  const provider = audit?.provider || {};
  const values = [provider.silentFallback, provider.silent_fallback, audit?.silentFallback, audit?.silent_fallback]
    .filter((value) => value !== undefined);
  return values.length > 0 && values.every((value) => value === false || value === "false");
}

function typedBoardWorkflow(commandText) {
  return /DocumentFile\.importDocx/i.test(commandText)
    && /DocumentFile\.patchDocx/i.test(commandText)
    && /DocumentFile\.importDocx/i.test(commandText)
    && /board-review-surgical-edit-workflow\.mjs/i.test(commandText);
}

export async function inspectBoardReviewDocx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const entries = await Promise.all(paths.map(async (part) => [part, Buffer.from(await zip.file(part).async("uint8array"))]));
  const partHashes = Object.fromEntries(entries.map(([part, value]) => [part, sha256(value)]));
  const content = new Map(entries.map(([part, value]) => [part, value.toString("utf8")]));
  const documentXml = content.get("word/document.xml") || "";
  const commentsXml = content.get("word/comments.xml") || "";
  const tables = topLevelXmlElements(documentXml, "tbl");
  const comments = parseComments(commentsXml);
  const modernParts = ["word/commentsExtended.xml", "word/commentsIds.xml", "word/commentsExtensible.xml", "word/people.xml"];
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes,
    documentXml,
    commentsXml,
    paragraphs: parseParagraphs(documentXml),
    tables: tables.map((table) => ({ text: wordText(table), rows: tableRows(table) })),
    comments,
    modern: {
      paths: modernParts.filter((part) => paths.includes(part)),
      hashes: Object.fromEntries(modernParts.filter((part) => partHashes[part]).map((part) => [part, partHashes[part]])),
      text: Object.fromEntries(modernParts.filter((part) => content.has(part)).map((part) => [part, content.get(part)])),
    },
  };
}

function findCommentAnchor(document, commentId, expectedText) {
  return document.paragraphs.find((paragraph) => paragraph.text === expectedText
    && paragraph.commentIds.filter((id) => String(id) === String(commentId)).length >= 3);
}

function visualChecks(visual) {
  const source = visual?.source;
  const output = visual?.output;
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount;
  const dimensionsStable = pageCountsMatch && source?.pages?.every((page, index) => {
    const next = output.pages?.[index];
    return next && page.width === next.width && page.height === next.height;
  });
  const changedPage = pageCountsMatch && source?.pages?.some((page, index) => page.pixelSha256 !== output.pages?.[index]?.pixelSha256);
  return { available, rendered, pageCountsMatch, dimensionsStable, changedPage };
}

export function gradeDocxBoardReviewEvidence({ evidence, audit, commands, item }) {
  const source = evidence.source;
  const output = evidence.output;
  const fixture = DOCX_SURGICAL_BOARD_REVIEW_FIXTURE;
  const sourceComment = source.comments.find((comment) => comment.author === fixture.comment.author);
  const outputComment = output.comments.find((comment) => String(comment.id) === String(sourceComment?.id));
  const sourceAnchor = findCommentAnchor(source, sourceComment?.id, fixture.recommendation.originalText);
  const outputAnchor = findCommentAnchor(output, outputComment?.id, fixture.recommendation.replacementText);
  const sourceRisk = source.tables.find((table) => table.rows.some((row) => row[0] === fixture.riskTable.targetLabel && row[1] === fixture.riskTable.originalStatus));
  const outputRisk = output.tables.find((table) => table.rows.some((row) => row[0] === fixture.riskTable.targetLabel && row[1] === fixture.riskTable.replacementStatus));
  const parts = changedParts(source, output);
  const modernStable = Object.keys(source.modern.hashes).length === 4
    && Object.entries(source.modern.hashes).every(([part, hash]) => output.modern.hashes[part] === hash)
    && Object.entries(source.modern.text).every(([part, text]) => output.modern.text[part] === text);
  const expectedDocument = source.documentXml
    .replace("Recommendation: continue the pilot.", "Recommendation: approve controlled release.")
    .replace(">Amber<", ">Green<");
  const expectedComments = source.commentsXml.replace("Please confirm the final retention wording.", "Confirmed by the audit committee.");
  const exactTargetParts = output.documentXml === expectedDocument && output.commentsXml === expectedComments;
  const visual = visualChecks(evidence.visual);
  const commandText = commands.join("\n");
  const outputEntries = evidence.outputEntries || [];
  return [
    check("docx-board-machine:recommendation-edit", "machine", Boolean(sourceAnchor && outputAnchor)
      && source.paragraphs.some((paragraph) => paragraph.text === fixture.recommendation.originalText)
      && output.paragraphs.some((paragraph) => paragraph.text === fixture.recommendation.replacementText), { sourceAnchor, outputAnchor }),
    check("docx-board-machine:risk-cell-edit", "machine", Boolean(sourceRisk && outputRisk), { sourceRisk, outputRisk }),
    check("docx-board-machine:classic-comment-edit", "machine", Boolean(sourceComment && outputComment)
      && sourceComment.author === outputComment.author
      && sourceComment.initials === outputComment.initials
      && sourceComment.date === outputComment.date
      && sourceComment.text === fixture.comment.originalText
      && outputComment.text === fixture.comment.replacementText, { sourceComment, outputComment }),
    check("docx-board-machine:reimport", "machine", audit?.validation?.reimport?.ok === true),
    check("docx-board-machine:all-preservation-canaries-present", "machine", [
      "word/footnotes.xml",
      "word/styles.xml",
      "word/commentsExtended.xml",
      "word/commentsIds.xml",
      "word/commentsExtensible.xml",
      "word/people.xml",
    ].every((part) => source.paths.includes(part) && output.paths.includes(part))),
    check("docx-board-machine:exact-target-xml", "machine", exactTargetParts, { parts, expectedChangedParts: ["word/comments.xml", "word/document.xml"] }),
    check("docx-board-machine:verification", "machine", audit?.validation?.verify?.ok === true),
    check("docx-board-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch && visual.dimensionsStable, { visual: evidence.visual }),
    check("docx-board-visual:target-change-visible", "visual", visual.changedPage === true, { visual: evidence.visual }),
    gate("docx-board-security:only-target-parts-changed", "security", parts.length === 2 && parts.includes("word/document.xml") && parts.includes("word/comments.xml"), { parts }),
    gate("docx-board-security:modern-graph-byte-stable", "security", modernStable, { source: source.modern.hashes, output: output.modern.hashes }),
    gate("docx-board-security:source-bound-audit", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256
      && audit?.provider?.silentFallback === false, { source: source.sha256, output: output.sha256, audit }),
    check("docx-board-security:source-immutable", "security", auditHash(audit, "source") === source.sha256),
    check("docx-board-trace:office-kit-provider", "trace", /office[- ]?kit/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), { provider: auditProvider(audit), version: auditVersion(audit) }),
    gate("docx-board-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("docx-board-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), { strategy: auditStrategy(audit) }),
    check("docx-board-trace:typed-primitive", "trace", typedBoardWorkflow(commandText), { commands }),
    check("docx-board-trace:operation", "trace", /surgical-board-review-edit/i.test(auditOperation(audit)), { operation: auditOperation(audit) }),
    check("docx-board-trace:output-contract", "trace", outputEntries.includes("board-review-updated.docx") && outputEntries.includes("audit.json") && outputEntries.length === 2, { outputEntries }),
  ];
}
