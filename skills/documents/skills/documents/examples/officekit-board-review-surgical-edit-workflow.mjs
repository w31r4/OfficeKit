import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";

import { DocumentFile, FileBlob } from "office-kit";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

function decodeXml(value = "") {
  return String(value)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function escapeXml(value = "") {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function wordText(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
    .map((match) => decodeXml(match[1].replace(/<[^>]+>/g, "")))
    .join("");
}

function xmlAttributes(opening = "") {
  const result = {};
  for (const match of String(opening).matchAll(/([:\w.-]+)="([^"]*)"/g)) {
    result[match[1].split(":").at(-1)] = decodeXml(match[2]);
  }
  return result;
}

function comments(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?comment\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?comment>/g)].map((match) => ({
    ...xmlAttributes(/^<(?:[\w.-]+:)?comment\b[^>]*>/.exec(match[0])?.[0] || ""),
    text: wordText(match[1]),
  }));
}

function paragraphs(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?p>/g)].map((match) => ({
    text: wordText(match[1]),
    commentIds: [...match[1].matchAll(/<(?:[\w.-]+:)?comment(?:RangeStart|RangeEnd|Reference)\b[^>]*>/g)]
      .map((entry) => xmlAttributes(entry[0]).id)
      .filter(Boolean),
  }));
}

function topLevelTables(xml = "") {
  const tables = [];
  let depth = 0;
  let start = -1;
  for (const token of String(xml).matchAll(/<\/?(?:[\w.-]+:)?tbl\b[^>]*>/g)) {
    if (/^<\//.test(token[0])) {
      depth -= 1;
      if (depth === 0 && start >= 0) {
        tables.push(String(xml).slice(start, (token.index || 0) + token[0].length));
        start = -1;
      }
    } else {
      if (depth === 0) start = token.index || 0;
      depth += 1;
    }
  }
  return tables;
}

async function packageVersion() {
  const entry = require.resolve("office-kit");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

async function inspectBoardSource(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const read = async (name) => zip.file(name)?.async("text") || "";
  const documentXml = await read("word/document.xml");
  const commentsXml = await read("word/comments.xml");
  const commentsExtendedXml = await read("word/commentsExtended.xml");
  const commentsIdsXml = await read("word/commentsIds.xml");
  const commentsExtensibleXml = await read("word/commentsExtensible.xml");
  const peopleXml = await read("word/people.xml");
  const sourceComments = comments(commentsXml);
  const targetComments = sourceComments.filter((comment) => comment.author === "Audit committee");
  if (targetComments.length !== 1) throw new Error("Expected exactly one Audit committee classic comment.");
  const targetComment = targetComments[0];
  if (targetComment.text !== "Please confirm the final retention wording.") {
    throw new Error("The board-review classic comment precondition did not match.");
  }
  const targetParagraphs = paragraphs(documentXml).filter((paragraph) => paragraph.text === "Recommendation: continue the pilot."
    && paragraph.commentIds.includes(String(targetComment.id)));
  if (targetParagraphs.length !== 1 || !targetParagraphs[0].commentIds.includes(String(targetComment.id))) {
    throw new Error("The board-review recommendation/comment anchor is not unique or is not bound to the expected comment.");
  }
  const targetTables = topLevelTables(documentXml).filter((table) => table.includes("Data migration") && table.includes("Amber"));
  if (targetTables.length !== 1) throw new Error("Expected exactly one risk table containing Data migration/Amber.");
  const modernParts = [commentsExtendedXml, commentsIdsXml, commentsExtensibleXml, peopleXml];
  if (modernParts.some((part) => !part)) throw new Error("The board-review modern comment identity graph is incomplete.");
  for (const paraId of ["17777777", "18888888", "19999999"]) {
    if (!new RegExp(`paraId=\\"${paraId}\\"`).test(commentsExtendedXml)
      || !new RegExp(`paraId=\\"${paraId}\\"`).test(commentsIdsXml)) {
      throw new Error(`The board-review modern comment graph is missing paraId ${paraId}.`);
    }
  }
  return {
    documentXml,
    commentsXml,
    targetCommentId: String(targetComment.id),
    targetParagraphText: targetParagraphs[0].text,
    targetTableText: wordText(targetTables[0]),
  };
}

function replaceOnce(value, search, replacement, label) {
  const index = value.indexOf(search);
  if (index < 0 || value.indexOf(search, index + search.length) >= 0) {
    throw new Error(`${label} must occur exactly once before mutation.`);
  }
  return value.slice(0, index) + replacement + value.slice(index + search.length);
}

function replaceFirst(value, search, replacement, label) {
  const index = value.indexOf(search);
  if (index < 0) throw new Error(`${label} was not found before mutation.`);
  return value.slice(0, index) + replacement + value.slice(index + search.length);
}

async function renderModel(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

export async function editBoardReview({ inputPath, outputPath, auditPath }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("input, output, and audit paths must be distinct so the source remains immutable.");
  }
  const source = await fs.readFile(sourcePath);
  const inspected = await inspectBoardSource(source);
  const imported = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const recommendation = imported.blocks.find((block) => block.kind === "paragraph" && block.text === inspected.targetParagraphText);
  const riskTable = imported.blocks.find((block) => block.kind === "table"
    && block.values?.some((row) => row?.[0] === "Data migration" && row?.[1] === "Amber"));
  if (!recommendation || recommendation.sourceBound === false || recommendation.textPatchable === false) {
    throw new Error("The recommendation paragraph is not an imported source-bound patch target.");
  }
  if (!riskTable) throw new Error("The imported risk table target is missing.");
  const tableTarget = riskTable.cells?.find((cell) => cell.row === 1 && cell.column === 1)
    || riskTable.cells?.find((cell) => riskTable.values?.[cell.row]?.[cell.column] === "Amber");
  if (!tableTarget || tableTarget.textPatchable === false) throw new Error("The risk status cell is not an editable typed target.");
  const outputDocumentXml = replaceFirst(
    replaceOnce(inspected.documentXml, ">Amber<", ">Green<", "risk status"),
    "Recommendation: continue the pilot.",
    "Recommendation: approve controlled release.",
    "recommendation",
  );
  const outputCommentsXml = replaceOnce(
    inspected.commentsXml,
    "Please confirm the final retention wording.",
    "Confirmed by the audit committee.",
    "classic comment",
  );
  const patched = await DocumentFile.patchDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }), [
    { path: "word/document.xml", xml: outputDocumentXml },
    { path: "word/comments.xml", xml: outputCommentsXml },
  ]);
  const output = new Uint8Array(await patched.arrayBuffer());
  const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
  const reimportedRecommendation = reimported.blocks.find((block) => block.kind === "paragraph" && block.text === "Recommendation: approve controlled release.");
  const reimportedRiskTable = reimported.blocks.find((block) => block.kind === "table"
    && block.values?.some((row) => row?.[0] === "Data migration" && row?.[1] === "Green"));
  if (!reimportedRecommendation || !reimportedRiskTable) throw new Error("Board-review output did not reimport the two requested semantic edits.");
  const outputInspection = await inspectBoardSource(output).catch(async (error) => {
    // The source precondition intentionally expects the old values.  Re-read
    // the output with the same independent package reader while checking all
    // identity/topology canaries explicitly.
    const zip = await JSZip.loadAsync(output);
    const read = async (name) => zip.file(name)?.async("text") || "";
    const documentXml = await read("word/document.xml");
    const commentsXml = await read("word/comments.xml");
    const commentsExtendedXml = await read("word/commentsExtended.xml");
    const commentsIdsXml = await read("word/commentsIds.xml");
    const commentsExtensibleXml = await read("word/commentsExtensible.xml");
    const peopleXml = await read("word/people.xml");
    const outputComments = comments(commentsXml);
    const target = outputComments.find((comment) => String(comment.id) === inspected.targetCommentId);
    if (!target || target.author !== "Audit committee" || target.text !== "Confirmed by the audit committee.") throw error;
    if (commentsExtendedXml === "" || commentsIdsXml === "" || commentsExtensibleXml === "" || peopleXml === "") throw error;
    for (const paraId of ["17777777", "18888888", "19999999"]) {
      if (!new RegExp(`paraId=\\"${paraId}\\"`).test(commentsExtendedXml) || !new RegExp(`paraId=\\"${paraId}\\"`).test(commentsIdsXml)) throw error;
    }
    return { documentXml, commentsXml, targetCommentId: inspected.targetCommentId };
  });
  const verification = reimported.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Document verification failed: " + verification.ndjson);
  const modelRender = await renderModel(reimported);
  const sourceZip = await JSZip.loadAsync(source);
  const outputZip = await JSZip.loadAsync(output);
  const paths = [...new Set([...Object.keys(sourceZip.files), ...Object.keys(outputZip.files)])].filter((entry) => !entry.endsWith("/")).sort();
  const changedParts = [];
  for (const part of paths) {
    const left = sourceZip.file(part) ? Buffer.from(await sourceZip.file(part).async("uint8array")) : null;
    const right = outputZip.file(part) ? Buffer.from(await outputZip.file(part).async("uint8array")) : null;
    if (!left || !right || !left.equals(right)) changedParts.push(part);
  }
  if (changedParts.join("|") !== "word/comments.xml|word/document.xml") {
    throw new Error(`Board-review edit changed unexpected package parts: ${changedParts.join(", ")}`);
  }
  const audit = {
    schema: "office-kit.docx-audit.v1",
    status: "succeeded",
    source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
    output: { path: finalPath, sha256: sha256(output), bytes: output.length },
    provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    operation: {
      type: "surgical-board-review-edit",
      targetParagraph: "Recommendation: continue the pilot.",
      targetTableCell: { row: 1, column: 1, from: "Amber", to: "Green" },
      commentId: inspected.targetCommentId,
      changedParts,
    },
    preservationOnly: [
      "word/commentsExtended.xml",
      "word/commentsIds.xml",
      "word/commentsExtensible.xml",
      "word/people.xml",
      "word/footnotes.xml",
      "word/styles.xml",
      "word/header*.xml",
      "word/footer*.xml",
      "TOC/REF fields",
      "content controls",
      "tracked revisions",
      "watermark",
    ],
    warnings: ["Mixed classic/modern comment graphs remain source-bound; only the uniquely bound classic comment text was edited."],
    validation: {
      reimport: { ok: true, recommendation: reimportedRecommendation.text, riskStatus: "Green", classicComment: outputInspection.targetCommentId },
      verify: { ok: verification.ok },
      modelRender: { ok: true, ...modelRender },
      changedParts,
    },
  };
  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    await fs.access(finalPath).then(() => { throw new Error("Refusing to overwrite an existing board-review output."); }, () => {});
    await fs.access(finalAuditPath).then(() => { throw new Error("Refusing to overwrite an existing board-review audit."); }, () => {});
    await fs.writeFile(temporaryPath, output, { flag: "wx" });
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2), { flag: "wx" });
    await fs.rename(temporaryPath, finalPath);
    await fs.rename(temporaryAuditPath, finalAuditPath);
  } catch (error) {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
  return { outputPath: finalPath, auditPath: finalAuditPath, audit };
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath] = argv;
  return { inputPath, outputPath, auditPath };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editBoardReview(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256 }));
}
