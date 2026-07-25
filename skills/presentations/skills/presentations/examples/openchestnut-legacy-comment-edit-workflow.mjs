import { constants as fsConstants } from "node:fs";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "open-office-artifact-tool";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_COMMENT_TEXT_LENGTH = 8_192;
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  return JSON.parse(await fs.readFile(path.join(path.dirname(path.dirname(entry)), "package.json"), "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty string.");
  return value.trim();
}

function commentText(value, label) {
  if (typeof value !== "string" || value.length > MAX_COMMENT_TEXT_LENGTH) {
    throw new TypeError(label + " must be a string no longer than " + MAX_COMMENT_TEXT_LENGTH + " characters.");
  }
  return value;
}

function snapshot(value) {
  return JSON.parse(JSON.stringify(value));
}

function canonical(value) {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === "object") return Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonical(value[key])]));
  return value;
}

function sameValue(left, right) {
  return JSON.stringify(canonical(left)) === JSON.stringify(canonical(right));
}

function xmlAttributes(tag) {
  const attributes = Object.create(null);
  for (const match of tag.matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g)) attributes[match[1]] = match[3];
  return attributes;
}

function relationshipPartPath(partPath) {
  const slash = partPath.lastIndexOf("/");
  const directory = slash < 0 ? "" : partPath.slice(0, slash + 1);
  const name = slash < 0 ? partPath : partPath.slice(slash + 1);
  return directory + "_rels/" + name + ".rels";
}

function resolveRelationshipTarget(sourcePart, target) {
  const resolved = new URL(target, "https://openchestnut.invalid/" + sourcePart);
  if (resolved.origin !== "https://openchestnut.invalid") throw new Error("Unexpected external PPTX relationship target.");
  const partPath = resolved.pathname.replace(/^\/+/, "");
  if (!partPath.startsWith("ppt/") || partPath.split("/").includes("..")) throw new Error("Unsafe PPTX relationship target.");
  return partPath;
}

async function relationshipEntries(zip, sourcePart) {
  const xml = await zip.file(relationshipPartPath(sourcePart))?.async("text");
  if (!xml) return [];
  return [...xml.matchAll(/<Relationship\b[^>]*>/gi)].map((match) => {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type || !attributes.Target) throw new Error("Malformed relationship for " + sourcePart + ".");
    const external = attributes.TargetMode?.toLowerCase() === "external";
    return {
      id: attributes.Id,
      type: attributes.Type,
      external,
      targetPart: external ? null : resolveRelationshipTarget(sourcePart, attributes.Target),
    };
  });
}

function relationshipType(entry, suffix) {
  return entry.type.endsWith("/" + suffix);
}

function exactlyOne(entries, suffix, label) {
  const matches = entries.filter((entry) => relationshipType(entry, suffix));
  if (matches.length !== 1 || matches[0].external || !matches[0].targetPart) {
    throw new Error(label + " must own exactly one internal " + suffix + " relationship.");
  }
  return matches[0];
}

async function orderedSlidePartPaths(zip) {
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  if (!presentationXml) throw new Error("PPTX is missing ppt/presentation.xml.");
  const relationships = await relationshipEntries(zip, "ppt/presentation.xml");
  const slideById = new Map(relationships
    .filter((entry) => relationshipType(entry, "slide") && !entry.external)
    .map((entry) => [entry.id, entry.targetPart]));
  const paths = [...presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*>/gi)].map((match) => {
    const partPath = slideById.get(xmlAttributes(match[0])["r:id"]);
    if (!partPath || !zip.file(partPath)) throw new Error("Presentation slide list contains an unresolved SlidePart.");
    return partPath;
  });
  if (!paths.length || new Set(paths).size !== paths.length) throw new Error("Presentation slide list must contain distinct SlideParts.");
  return paths;
}

async function entryHashes(zip) {
  const hashes = new Map();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (!entry.dir) hashes.set(name, sha256(await entry.async("uint8array")));
  }
  return hashes;
}

async function auditTextOnlyLegacyCommentEdit(sourceBytes, outputBytes, targetIndex) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const sourceSlides = await orderedSlidePartPaths(sourceZip);
  const outputSlides = await orderedSlidePartPaths(outputZip);
  if (!sameValue(sourceSlides, outputSlides)) throw new Error("Legacy comment edit changed SlidePart identity or order.");
  const targetSlidePart = sourceSlides[targetIndex];
  if (!targetSlidePart) throw new Error("Target slide index is outside the source package slide list.");

  const sourcePresentationRelationships = await relationshipEntries(sourceZip, "ppt/presentation.xml");
  const outputPresentationRelationships = await relationshipEntries(outputZip, "ppt/presentation.xml");
  const sourceAuthors = exactlyOne(sourcePresentationRelationships, "commentAuthors", "Source PresentationPart");
  const outputAuthors = exactlyOne(outputPresentationRelationships, "commentAuthors", "Output PresentationPart");
  if (!sameValue(sourceAuthors, outputAuthors) || sourceAuthors.targetPart !== "ppt/commentAuthors.xml") {
    throw new Error("Legacy comment edit changed the shared author-catalog relationship.");
  }
  if ((await relationshipEntries(sourceZip, sourceAuthors.targetPart)).length ||
      (await relationshipEntries(outputZip, outputAuthors.targetPart)).length) {
    throw new Error("Legacy CommentAuthorsPart must remain a closed leaf.");
  }

  const sourceComments = exactlyOne(await relationshipEntries(sourceZip, targetSlidePart), "comments", "Source target SlidePart");
  const outputComments = exactlyOne(await relationshipEntries(outputZip, targetSlidePart), "comments", "Output target SlidePart");
  if (!sameValue(sourceComments, outputComments) || !/^ppt\/comments\/comment\d+\.xml$/.test(sourceComments.targetPart)) {
    throw new Error("Legacy comment edit changed the target SlideCommentsPart relationship or naming profile.");
  }
  if ((await relationshipEntries(sourceZip, sourceComments.targetPart)).length ||
      (await relationshipEntries(outputZip, outputComments.targetPart)).length) {
    throw new Error("Legacy SlideCommentsPart must remain a closed leaf.");
  }

  const before = await entryHashes(sourceZip);
  const after = await entryHashes(outputZip);
  if (!sameValue([...before.keys()].sort(), [...after.keys()].sort())) throw new Error("Legacy comment edit added or removed an OPC package part.");
  const changed = [...before.keys()].filter((name) => before.get(name) !== after.get(name)).sort();
  const expectedChanged = [sourceComments.targetPart];
  if (!sameValue(changed, expectedChanged)) {
    throw new Error("Legacy comment text edit changed unexpected package parts: " + changed.join(", "));
  }
  return {
    ok: true,
    targetSlidePart,
    authorsPart: sourceAuthors.targetPart,
    commentsPart: sourceComments.targetPart,
    changedParts: changed,
  };
}

function visibleSlideSnapshot(slide) {
  const value = structuredClone(slide.toProto());
  delete value.comments;
  return JSON.stringify(value);
}

async function modelSvg(slide) {
  const rendered = await slide.export({ format: "svg" });
  if (!/<svg\b/i.test(await rendered.text())) throw new Error("Presentation model render did not produce SVG.");
  return { bytes: rendered.bytes.length, sha256: sha256(rendered.bytes) };
}

async function copyExclusive(source, destination) {
  await fs.copyFile(source, destination, fsConstants.COPYFILE_EXCL);
}

function resolveTarget(presentation, { slideName, commentId, expectedText }) {
  const slides = presentation.slides.items.filter((slide) => slide.name === slideName);
  if (slides.length !== 1) throw new Error("Expected exactly one imported slide named " + JSON.stringify(slideName) + "; found " + slides.length + ".");
  const slide = slides[0];
  const capability = snapshot(slide.comments.capability);
  if (!capability?.sourceBound || capability.format !== "legacy" || !capability.partPresent || !capability.editable || capability.addable) {
    throw new Error("Selected slide does not expose the canonical imported legacy-comment text-edit capability.");
  }
  const matches = slide.comments.items.filter((thread) => thread.id === commentId);
  if (matches.length !== 1) throw new Error("Expected exactly one imported legacy comment with ID " + JSON.stringify(commentId) + "; found " + matches.length + ".");
  const thread = matches[0];
  if (thread.nativeFormat !== "legacy" || thread.targetId || thread.resolved || !Array.isArray(thread.comments) || thread.comments.length !== 1 ||
      thread.comments[0].text !== expectedText || typeof thread.comments[0].author !== "string" || !thread.nativeAnchor?.format) {
    throw new Error("Selected legacy comment does not match the fixed root-text source-bound profile.");
  }
  return { slide, thread, capability };
}

/**
 * Replace exactly one existing legacy PPTX review-comment root text.
 *
 * The operation is deliberately not a general comment editor: it requires the
 * imported legacy text-edit capability, a stable model ID, and exact old text.
 * It keeps author, created timestamp, coordinate, native author/index identity,
 * thread count/order, relationships, and every non-comment package part fixed.
 */
export async function editPptxLegacyReviewComment({ inputPath, outputPath, auditPath, slideName, commentId, expectedText, replacementText }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (new Set([sourcePath, finalPath, finalAuditPath]).size !== 3) throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  const config = {
    slideName: requiredText(slideName, "slideName"),
    commentId: requiredText(commentId, "commentId"),
    expectedText: commentText(expectedText, "expectedText"),
    replacementText: commentText(replacementText, "replacementText"),
  };
  if (config.expectedText === config.replacementText) throw new Error("replacementText must differ from expectedText so the transaction has a declared semantic change.");

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const originalSlideNames = presentation.slides.items.map((slide) => slide.name);
  const target = resolveTarget(presentation, config);
  const targetIndex = presentation.slides.items.indexOf(target.slide);
  const visibleSnapshot = visibleSlideSnapshot(target.slide);
  const sourceRender = await modelSvg(target.slide);
  const originalThreads = snapshot(target.slide.comments.items.map((thread) => thread.toJSON()));
  const expectedThreads = snapshot(originalThreads);
  const targetThreadIndex = target.slide.comments.items.indexOf(target.thread);
  expectedThreads[targetThreadIndex].comments[0].text = config.replacementText;
  target.thread.comments[0].text = config.replacementText;

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  let publishedOutput = false;
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageGraph = await auditTextOnlyLegacyCommentEdit(source, output, targetIndex);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    if (!sameValue(reimported.slides.items.map((slide) => slide.name), originalSlideNames)) throw new Error("PPTX export changed slide count, order, or names.");
    const roundTrip = resolveTarget(reimported, { ...config, expectedText: config.replacementText });
    if (visibleSlideSnapshot(roundTrip.slide) !== visibleSnapshot) throw new Error("Legacy review-comment edit changed visible target-slide semantics.");
    if (!sameValue(roundTrip.slide.comments.items.map((thread) => thread.toJSON()), expectedThreads)) {
      throw new Error("PPTX export changed legacy-comment author, timestamp, coordinate, native identity, order, or thread topology.");
    }
    const outputRender = await modelSvg(roundTrip.slide);
    if (outputRender.sha256 !== sourceRender.sha256) throw new Error("Editing a nonvisual legacy review comment changed the target slide model SVG.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source presentation changed during the transaction.");

    const audit = {
      schema: "open-office-artifact-tool.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", overwrite: false },
      operation: {
        type: "source-bound-legacy-comment-text-edit",
        slideId: target.slide.id,
        slideName: target.slide.name,
        slideIndex: targetIndex,
        commentId: config.commentId,
        expectedText: config.expectedText,
        replacementText: config.replacementText,
      },
      precondition: { capability: target.capability, originalComment: originalThreads[targetThreadIndex] },
      validation: {
        package: packageGraph,
        reimport: { ok: true, slideCount: reimported.slides.count, visibleSemanticsPreserved: true, fixedTopology: true, capability: roundTrip.capability },
        modelRender: { ok: true, renderer: "model-svg", sourceSha256: sourceRender.sha256, outputSha256: outputRender.sha256, byteIdentical: true },
        verify: { ok: verification.ok },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await copyExclusive(temporaryPath, finalPath);
    publishedOutput = true;
    await copyExclusive(temporaryAuditPath, finalAuditPath);
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    if (publishedOutput) await fs.rm(finalPath, { force: true });
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, slideName, commentId, expectedText, replacementText] = argv;
  return { inputPath, outputPath, auditPath, slideName, commentId, expectedText, replacementText };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxLegacyReviewComment(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256, commentsPart: result.audit.validation.package.commentsPart }));
}
