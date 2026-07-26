import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { DocumentFile, FileBlob } from "office-kit";
import {
  DOCX_MIME,
  assertAbsent,
  canonicalizeXmlForResidual,
  changedParts,
  packageVersion,
  publishNoReplace,
  readPackagePartText,
  requiredText,
  sha256,
} from "../artifact_tool/_source_bound_docx.mjs";

const NOTE_PARTS = Object.freeze({
  footnote: "word/footnotes.xml",
  endnote: "word/endnotes.xml",
});

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function noteKind(value, label) {
  if (value !== "footnote" && value !== "endnote") throw new TypeError(`${label} must be footnote or endnote.`);
  return value;
}

function boundedNativeId(value, label) {
  const nativeId = Number(value);
  if (!Number.isInteger(nativeId) || nativeId < 1 || nativeId > 2_147_483_647) {
    throw new TypeError(`${label} must be a positive 32-bit integer.`);
  }
  return nativeId;
}

function boundedParagraphIndex(value, label) {
  const paragraphIndex = Number(value);
  if (!Number.isSafeInteger(paragraphIndex) || paragraphIndex < 0 || paragraphIndex > 15) {
    throw new TypeError(`${label} must be an integer from 0 through 15.`);
  }
  return paragraphIndex;
}

function canonicalParagraphText(value, label) {
  const text = requiredText(value, label);
  if (text !== text.trim()) throw new TypeError(`${label} must not have leading or trailing whitespace in this source-bound workflow.`);
  if (/[\r\n\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(text)) {
    throw new TypeError(`${label} must be one XML-safe physical paragraph without line breaks or control characters.`);
  }
  return text;
}

function normalizeTarget(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("target must be an object.");
  const allowed = new Set(["kind", "noteId", "nativeId", "targetId", "paragraphIndex", "expectedText"]);
  const unsupported = Object.keys(value).filter((key) => !allowed.has(key));
  if (unsupported.length) throw new TypeError(`target has unsupported properties: ${unsupported.join(", ")}.`);
  return {
    kind: noteKind(value.kind, "target.kind"),
    noteId: requiredText(value.noteId, "target.noteId"),
    nativeId: boundedNativeId(value.nativeId, "target.nativeId"),
    targetId: requiredText(value.targetId, "target.targetId"),
    paragraphIndex: boundedParagraphIndex(value.paragraphIndex, "target.paragraphIndex"),
    expectedText: canonicalParagraphText(value.expectedText, "target.expectedText"),
  };
}

function noteSnapshot(note) {
  return {
    id: note.id,
    kind: note.kind,
    name: note.name || "",
    nativeId: note.nativeId,
    targetId: note.targetId,
    paragraphs: note.paragraphs,
  };
}

function documentProjection(document) {
  return {
    blocks: document.blocks.map((block) => ({ id: block.id, kind: block.kind, text: block.text })),
    notes: document.notes.map(noteSnapshot),
  };
}

function selectNote(document, target) {
  const matches = document.notes.filter((note) =>
    note.id === target.noteId
    && note.kind === target.kind
    && note.nativeId === target.nativeId
    && note.targetId === target.targetId);
  if (matches.length !== 1) {
    throw new Error(`Expected exactly one ${target.kind} matching the inspected note ID, native ID, and target ID; found ${matches.length}.`);
  }
  const note = matches[0];
  if (document.resolve(note.id) !== note) throw new Error("The selected note locator did not resolve to the inspected object.");
  const paragraphs = note.paragraphs;
  if (paragraphs.length < 1 || paragraphs.length > 16 || target.paragraphIndex >= paragraphs.length) {
    throw new Error("The selected note does not retain the fixed 1-through-16 paragraph profile required by this workflow.");
  }
  for (const [index, paragraph] of paragraphs.entries()) canonicalParagraphText(paragraph, `selected note paragraph ${index}`);
  if (paragraphs[target.paragraphIndex] !== target.expectedText) {
    throw new Error(`Selected note paragraph does not match the expected source text: expected ${JSON.stringify(target.expectedText)}, observed ${JSON.stringify(paragraphs[target.paragraphIndex])}.`);
  }
  return { note, snapshot: noteSnapshot(note), paragraphs };
}

function xmlAttributes(source = "") {
  const attributes = {};
  for (const match of String(source).matchAll(/([:\w.-]+)="([^"]*)"/g)) attributes[match[1]] = match[2];
  return attributes;
}

function decodeXmlText(value, label) {
  const source = String(value);
  let cursor = 0;
  let decoded = "";
  for (const match of source.matchAll(/&(?:amp|lt|gt|quot|apos|#\d+|#x[\da-fA-F]+);/g)) {
    const prefix = source.slice(cursor, match.index);
    if (prefix.includes("&")) throw new Error(`${label} has an unsupported XML entity.`);
    const entity = match[0];
    decoded += prefix;
    if (entity === "&amp;") decoded += "&";
    else if (entity === "&lt;") decoded += "<";
    else if (entity === "&gt;") decoded += ">";
    else if (entity === "&quot;") decoded += '"';
    else if (entity === "&apos;") decoded += "'";
    else {
      const numeric = entity.startsWith("&#x") ? Number.parseInt(entity.slice(3, -1), 16) : Number.parseInt(entity.slice(2, -1), 10);
      if (!Number.isInteger(numeric) || numeric < 0 || numeric > 0x10ffff) throw new Error(`${label} has an invalid XML character reference.`);
      decoded += String.fromCodePoint(numeric);
    }
    cursor = match.index + entity.length;
  }
  const suffix = source.slice(cursor);
  if (suffix.includes("&")) throw new Error(`${label} has an unsupported XML entity.`);
  return decoded + suffix;
}

function canonicalNoteParagraph(paragraphXml, kind, paragraphIndex, label) {
  const first = paragraphIndex === 0;
  const pattern = first
    ? new RegExp(`^<w:r><w:${kind}Ref\\s*/></w:r><w:r><w:t xml:space="preserve">([\\s\\S]*)</w:t></w:r>$`)
    : /^<w:r><w:t xml:space="preserve">([\s\S]*)<\/w:t><\/w:r>$/;
  const match = pattern.exec(paragraphXml);
  if (!match) {
    throw new Error(`${label} paragraph ${paragraphIndex + 1} is outside the canonical marker-plus-one-text-run profile.`);
  }
  const rawText = match[1];
  const textOffset = paragraphXml.length - "</w:t></w:r>".length - rawText.length;
  return { rawText, visibleText: decodeXmlText(rawText, `${label} paragraph ${paragraphIndex + 1}`), textOffset };
}

function canonicalNativeNote(xml, { kind, nativeId, paragraphs }, label) {
  const elementName = kind === "footnote" ? "footnote" : "endnote";
  const matches = [];
  const pattern = new RegExp(`<w:${elementName}\\b([^>]*)>([\\s\\S]*?)</w:${elementName}>`, "g");
  for (const match of String(xml).matchAll(pattern)) {
    const attributes = xmlAttributes(match[1]);
    if (attributes["w:id"] === String(nativeId)) matches.push({ match, attributes });
  }
  if (matches.length !== 1) throw new Error(`${label} must contain exactly one w:${elementName} with w:id=${nativeId}; found ${matches.length}.`);
  const { match, attributes } = matches[0];
  const unsupportedAttributes = Object.keys(attributes).filter((name) => name !== "w:id");
  if (unsupportedAttributes.length) throw new Error(`${label} has unsupported positive-note attributes: ${unsupportedAttributes.join(", ")}.`);
  const elementXml = match[0];
  const body = match[2];
  const paragraphMatches = [...body.matchAll(/<w:p\b([^>]*)>([\s\S]*?)<\/w:p>/g)];
  if (!paragraphMatches.length || paragraphMatches.map((item) => item[0]).join("") !== body) {
    throw new Error(`${label} is not a direct sequence of canonical w:p note paragraphs.`);
  }
  if (paragraphMatches.length !== paragraphs.length) {
    throw new Error(`${label} paragraph count does not match the imported model: expected ${paragraphs.length}, found ${paragraphMatches.length}.`);
  }
  const openingLength = elementXml.indexOf(">") + 1;
  const parsed = paragraphMatches.map((paragraphMatch, paragraphIndex) => {
    if (paragraphMatch[1]) throw new Error(`${label} paragraph ${paragraphIndex + 1} has unsupported w:p attributes.`);
    const canonical = canonicalNoteParagraph(paragraphMatch[2], kind, paragraphIndex, label);
    const expectedVisibleText = `${paragraphIndex === 0 ? " " : ""}${paragraphs[paragraphIndex]}`;
    if (canonical.visibleText !== expectedVisibleText) {
      throw new Error(`${label} paragraph ${paragraphIndex + 1} does not match the imported model text.`);
    }
    return {
      ...canonical,
      contentStart: match.index + openingLength + paragraphMatch.index + paragraphMatch[0].indexOf(">") + 1 + canonical.textOffset,
      contentEnd: match.index + openingLength + paragraphMatch.index + paragraphMatch[0].indexOf(">") + 1 + canonical.textOffset + canonical.rawText.length,
    };
  });
  return { paragraphs: parsed };
}

function maskNativeNoteParagraph(xml, native, paragraphIndex, label) {
  const paragraph = native.paragraphs[paragraphIndex];
  if (!paragraph) throw new Error(`${label} selected paragraph is absent from the native note.`);
  return canonicalizeXmlForResidual(
    `${String(xml).slice(0, paragraph.contentStart)}__OFFICE_KIT_TARGET_NOTE_TEXT__${String(xml).slice(paragraph.contentEnd)}`,
    label,
  );
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes exactly one physical paragraph of one recognized imported footnote or
 * endnote through the public DocumentFile path. It never moves a reference,
 * changes the note count, or rebuilds a general note graph.
 */
export async function editImportedNoteParagraphText({ inputPath, outputPath, auditPath, target, replacementText }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const selectedTarget = normalizeTarget(target);
  const replacement = canonicalParagraphText(replacementText, "replacementText");
  if (replacement === selectedTarget.expectedText) throw new Error("replacementText must differ from target.expectedText.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectNote(document, selectedTarget);
  const partPath = NOTE_PARTS[selectedTarget.kind];
  const sourceXml = await readPackagePartText(source, partPath, "Source DOCX package");
  const sourceNative = canonicalNativeNote(sourceXml, { ...selectedTarget, paragraphs: selected.paragraphs }, "Source target note");
  const sourceResidual = maskNativeNoteParagraph(sourceXml, sourceNative, selectedTarget.paragraphIndex, "Source target note");
  const beforeProjection = documentProjection(document);
  const nextParagraphs = [...selected.paragraphs];
  nextParagraphs[selectedTarget.paragraphIndex] = replacement;
  selected.note.paragraphs = nextParagraphs;

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound note paragraph edit");
    if (!equalJson(changed, [partPath])) {
      throw new Error(`Source-bound note paragraph edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }

    const outputXml = await readPackagePartText(output, partPath, "Output DOCX package");
    const outputNative = canonicalNativeNote(outputXml, { ...selectedTarget, paragraphs: nextParagraphs }, "Output target note");
    const outputResidual = maskNativeNoteParagraph(outputXml, outputNative, selectedTarget.paragraphIndex, "Output target note");
    if (outputResidual !== sourceResidual) {
      throw new Error(`${partPath} changed outside the one requested canonical note text payload.`);
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectNote(reimported, { ...selectedTarget, expectedText: replacement });
    const afterProjection = documentProjection(reimported);
    const expectedProjection = structuredClone(beforeProjection);
    const expectedNote = expectedProjection.notes.find((note) => note.id === selected.snapshot.id);
    if (!expectedNote) throw new Error("Selected note disappeared from the imported note projection.");
    expectedNote.paragraphs[selectedTarget.paragraphIndex] = replacement;
    if (!equalJson(afterProjection, expectedProjection)) {
      throw new Error("DOCX export changed body, note identity, anchor, or text outside the requested physical note paragraph.");
    }
    if (roundTrip.snapshot.id !== selected.snapshot.id || roundTrip.snapshot.nativeId !== selected.snapshot.nativeId || roundTrip.snapshot.targetId !== selected.snapshot.targetId) {
      throw new Error("Second import did not preserve the selected note identity, native ID, or anchor target.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
    const render = await modelRender(reimported);
    const audit = {
      schema: "office-kit.docx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", noReplace: true },
      operation: {
        type: "source-bound-note-paragraph-text-edit",
        target: {
          kind: selectedTarget.kind,
          noteId: selected.snapshot.id,
          nativeId: selected.snapshot.nativeId,
          targetId: selected.snapshot.targetId,
          paragraphIndex: selectedTarget.paragraphIndex,
        },
        sourceTextSha256: sha256(Buffer.from(selectedTarget.expectedText, "utf8")),
        replacementTextSha256: sha256(Buffer.from(replacement, "utf8")),
      },
      validation: {
        changedParts: changed,
        noteXmlResidual: {
          ok: true,
          partPath,
          paragraphIndex: selectedTarget.paragraphIndex,
          normalizedSha256: sha256(Buffer.from(sourceResidual, "utf8")),
        },
        reimport: {
          ok: true,
          noteId: roundTrip.snapshot.id,
          nativeId: roundTrip.snapshot.nativeId,
          targetId: roundTrip.snapshot.targetId,
          paragraphCount: roundTrip.snapshot.paragraphs.length,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes one canonical note text leaf only. Note anchors, counts, native IDs, paragraph topology, and formatting remain source-owned; render every affected page before delivery."],
    };
    await fs.writeFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`, { flag: "wx" });
    await publishNoReplace(temporaryPath, finalPath);
    try {
      await publishNoReplace(temporaryAuditPath, finalAuditPath);
    } catch (error) {
      await fs.rm(finalPath, { force: true });
      throw error;
    }
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

function parseJsonTarget(value) {
  try {
    return JSON.parse(requiredText(value, "target"));
  } catch (error) {
    throw new TypeError(`target must be JSON for a note locator object: ${error.message}`);
  }
}

export function parseNoteTextEditCli(argv) {
  const [inputPath, outputPath, auditPath, target, replacementText] = argv;
  return { inputPath, outputPath, auditPath, target: parseJsonTarget(target), replacementText };
}

export function noteTextCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedNoteParagraphText(parseNoteTextEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(noteTextCliOutput(result)));
}
