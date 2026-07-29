import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { FileBlob, PresentationFile } from "office-kit";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value;
}

function positiveInteger(value, label) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1) throw new TypeError(`${label} must be a positive integer.`);
  return parsed;
}

function sourceBlob(bytes, name) {
  return new FileBlob(bytes, { type: PPTX_MIME, name });
}

function parseNdjson(ndjson, label) {
  if (typeof ndjson !== "string") throw new Error(`${label} did not produce NDJSON.`);
  return ndjson.split("\n").filter(Boolean).map((line, index) => {
    try {
      return JSON.parse(line);
    } catch (error) {
      throw new Error(`${label} emitted invalid NDJSON record ${index + 1}: ${error.message}`);
    }
  });
}

async function packageVersion() {
  const entry = require.resolve("office-kit");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

async function assertAbsent(target, label) {
  try {
    await fs.lstat(target);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists: ${target}`);
}

function connectedDiagramCandidate(presentation, { smartArtName, slide }) {
  const slideModel = presentation.slides.items[slide - 1];
  if (!slideModel) throw new Error(`Slide ${slide} does not exist.`);
  const candidates = slideModel.nativeObjects.items
    .filter((object) => object.nativeKind === "diagram" && object.name === smartArtName);
  if (candidates.length !== 1) {
    throw new Error(`Expected exactly one named SmartArt ${JSON.stringify(smartArtName)} on slide ${slide}; found ${candidates.length}.`);
  }
  const object = candidates[0];
  if (object.editable !== false) {
    throw new Error("The selected SmartArt is editable; this audit-only refusal workflow must not replace an available typed edit path.");
  }
  const dataPart = object.parts.find((part) => Array.isArray(part.relationships)
    && part.relationships.some((relationship) => String(relationship.targetMode || "").toLowerCase() === "external"));
  if (!dataPart) {
    throw new Error("The selected SmartArt is not a connected source-bound graph with an external child relationship.");
  }
  return { slideModel, object, dataPart };
}

function expectedDiagramNode(dataPartRecord, { nodeModelId, expectedText }) {
  const preview = String(dataPartRecord?.textPreview || "");
  if (!preview.includes(nodeModelId)) throw new Error(`The connected DiagramDataPart does not contain model ID ${JSON.stringify(nodeModelId)}.`);
  if (!preview.includes(expectedText)) throw new Error(`The connected DiagramDataPart does not contain expected text ${JSON.stringify(expectedText)}.`);
  const nodeCount = (preview.match(/<dgm:pt\b/g) || []).length;
  if (!nodeCount) throw new Error("The connected DiagramDataPart does not expose any canonical dgm:pt nodes through public inspection.");
  return { nodeCount, preview };
}

function reviewCanaries(presentation, { notesSlide, modernCommentThread, expectedDirectReplies }) {
  const slide = presentation.slides.items[notesSlide - 1];
  if (!slide) throw new Error(`Notes/comment slide ${notesSlide} does not exist.`);
  if (!slide.speakerNotes?.capability?.partPresent) {
    throw new Error(`Slide ${notesSlide} has no existing source-bound speaker notes part.`);
  }
  const thread = slide.comments.items.find((candidate) => candidate.id === modernCommentThread);
  if (!thread || thread.nativeFormat !== "modern") {
    throw new Error(`Slide ${notesSlide} does not contain the expected modern comment thread ${JSON.stringify(modernCommentThread)}.`);
  }
  const directReplyCount = Math.max(0, thread.comments.length - 1);
  if (directReplyCount !== expectedDirectReplies) {
    throw new Error(`Modern comment thread ${JSON.stringify(modernCommentThread)} has ${directReplyCount} direct replies; expected ${expectedDirectReplies}.`);
  }
  return { thread, directReplyCount };
}

/**
 * Audit a compound request that includes an imported connected SmartArt graph.
 *
 * This workflow intentionally never exports a PPTX. It proves the named graph
 * and independent notes/comment canaries through public OfficeKit APIs, then
 * writes one canonical failed-closed audit. It is appropriate only when the
 * caller has already decided that all requested mutations are one transaction.
 */
export async function refuseConnectedSmartArtTransaction({
  inputPath,
  auditPath,
  smartArtName,
  nodeModelId,
  expectedText,
  slide,
  notesSlide,
  modernCommentThread,
  expectedDirectReplies = 1,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const target = {
    smartArtName: requiredText(smartArtName, "smartArtName"),
    nodeModelId: requiredText(nodeModelId, "nodeModelId"),
    expectedText: requiredText(expectedText, "expectedText"),
    slide: positiveInteger(slide, "slide"),
    notesSlide: positiveInteger(notesSlide, "notesSlide"),
    modernCommentThread: requiredText(modernCommentThread, "modernCommentThread"),
  };
  const replyCount = positiveInteger(expectedDirectReplies, "expectedDirectReplies");
  if (sourcePath === finalAuditPath) throw new Error("inputPath and auditPath must be distinct so the source remains immutable.");
  await assertAbsent(finalAuditPath, "auditPath");

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const blob = sourceBlob(source, path.basename(sourcePath));
  const presentation = await PresentationFile.importPptx(blob);
  // Keep the semantic inspect call explicit: the lower-level package inspect
  // below is used only to bind opaque DiagramDataPart node identity.
  const semanticInspection = presentation.inspect({ kind: "slide,nativeObject,diagram,notes,comment,thread", maxChars: 50_000 });
  if (typeof semanticInspection?.ndjson !== "string") throw new Error("Presentation semantic inspection did not return NDJSON.");

  const connected = connectedDiagramCandidate(presentation, target);
  const packageInspection = await PresentationFile.inspectPptx(blob, { includeText: true, maxPreviewChars: 12_000 });
  if (!packageInspection.ok) throw new Error(`PPTX package inspection failed: ${packageInspection.ndjson}`);
  const packageRecords = parseNdjson(packageInspection.ndjson, "PPTX package inspection");
  const dataPartRecord = packageRecords.find((record) => record.path === connected.dataPart.path);
  if (!dataPartRecord) throw new Error(`Public package inspection did not return ${connected.dataPart.path}.`);
  const node = expectedDiagramNode(dataPartRecord, target);
  const review = reviewCanaries(presentation, { ...target, expectedDirectReplies: replyCount });
  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Presentation verification failed before refusing the atomic transaction: ${verification.ndjson}`);

  const sourceAfterPreflight = await fs.readFile(sourcePath);
  if (!sourceAfterPreflight.equals(source)) throw new Error("Source bytes changed during preflight; no audit was written.");
  const audit = {
    schema: "office-kit.pptx-connected-smartart-refusal.v1",
    status: "failed_closed",
    source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
    output: null,
    officeKit: {
      actualProvider: "office-kit",
      actualVersion: await packageVersion(),
      silentFallback: false,
    },
    target: {
      smartArtName: target.smartArtName,
      nodeModelId: target.nodeModelId,
      slide: target.slide,
      notesSlide: target.notesSlide,
      modernCommentThread: target.modernCommentThread,
    },
    mutationAttempted: false,
    preflight: {
      import: { ok: true, method: "PresentationFile.importPptx" },
      inspect: {
        ok: true,
        method: "Presentation.inspect + PresentationFile.inspectPptx",
        slideCount: presentation.slides.items.length,
        smartArt: {
          name: target.smartArtName,
          nodeCount: node.nodeCount,
          connectedDataRelationship: true,
          editable: false,
        },
        review: {
          slide: target.notesSlide,
          speakerNotes: true,
          modernCommentThreadCount: presentation.slides.items[target.notesSlide - 1].comments.items.length,
          modernCommentReplyCount: review.directReplyCount,
        },
      },
      capabilityDecision: {
        supported: false,
        reason: "The selected DiagramDataPart owns an external child relationship, so OfficeKit does not expose a bounded SmartArt node-text mutation for this connected graph.",
      },
      verify: { ok: true, method: "Presentation.verify({ visualQa: true })" },
    },
    savePolicy: {
      strategy: "no-output-fail-closed",
      sourcePreserved: true,
      artifactPublished: false,
    },
    warnings: [
      "No SmartArt, speaker-notes, or modern-comment mutation was attempted because the requested transaction is atomic.",
    ],
    validation: {
      sourceHashRecorded: true,
      sourcePreserved: true,
      outputAbsent: true,
      packageInspection: { ok: true, dataPart: connected.dataPart.path },
      semanticInspection: { ok: true },
    },
  };
  const directory = path.dirname(finalAuditPath);
  const temporaryAuditPath = path.join(directory, `.${path.basename(finalAuditPath)}.tmp-${process.pid}-${Date.now()}`);
  await fs.mkdir(directory, { recursive: true });
  try {
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await fs.rename(temporaryAuditPath, finalAuditPath);
  } catch (error) {
    await fs.rm(temporaryAuditPath, { force: true });
    throw error;
  }
  return { auditPath: finalAuditPath, audit };
}

function parseCli(argv) {
  const values = {};
  const allowed = new Set([
    "--input",
    "--audit",
    "--smartart-name",
    "--node-model-id",
    "--expected-text",
    "--slide",
    "--notes-slide",
    "--modern-comment-thread",
    "--expected-direct-replies",
  ]);
  for (let index = 0; index < argv.length; index += 2) {
    const flag = argv[index];
    const value = argv[index + 1];
    if (!allowed.has(flag) || value === undefined) throw new Error(`Expected a value for a known flag; received ${JSON.stringify(flag)}.`);
    values[flag.slice(2).replace(/-([a-z])/g, (_, letter) => letter.toUpperCase())] = value;
  }
  return {
    inputPath: values.input,
    auditPath: values.audit,
    smartArtName: values.smartartName,
    nodeModelId: values.nodeModelId,
    expectedText: values.expectedText,
    slide: values.slide,
    notesSlide: values.notesSlide,
    modernCommentThread: values.modernCommentThread,
    expectedDirectReplies: values.expectedDirectReplies || 1,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await refuseConnectedSmartArtTransaction(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    status: result.audit.status,
    auditPath: result.auditPath,
    sourceSha256: result.audit.source.sha256,
    output: result.audit.output,
  }));
}
