import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { DocumentFile, FileBlob } from "office-kit";
import {
  DOCX_MIME,
  assertAbsent,
  changedParts,
  packageVersion,
  publishNoReplace,
  readPackagePartText,
  requiredText,
  sha256,
} from "../artifact_tool/_source_bound_docx.mjs";

const MAX_ALT_TEXT_CHARS = 32_767;
const MOVABLE_NAMESPACES = new Map([
  ["xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
  ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
  ["xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
  ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
  ["xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
]);

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

function normalizeAltText(value, label) {
  if (typeof value !== "string" || !value.length || value.length > MAX_ALT_TEXT_CHARS || /[\u0000-\u001f\u007f]/.test(value)) {
    throw new TypeError(`${label} must be a non-empty alternative-text string of at most ${MAX_ALT_TEXT_CHARS} characters without controls.`);
  }
  return value;
}

function decodeCanonicalXmlAttribute(value, label) {
  const text = String(value);
  if (/[<&]/.test(text.replace(/&(amp|lt|gt|quot|apos);/g, ""))) {
    throw new Error(`${label} is not canonically XML escaped.`);
  }
  const decoded = text.replace(/&(amp|lt|gt|quot|apos);/g, (_match, entity) => ({
    amp: "&",
    lt: "<",
    gt: ">",
    quot: "\"",
    apos: "'",
  })[entity]);
  return decoded;
}

function encodeXmlAttribute(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function parseCanonicalSelfClosingElement(markup, qualifiedName, label) {
  const match = new RegExp(`^<${qualifiedName}((?:\\s+[:\\w.-]+="[^"]*")*)\\s*/>$`).exec(String(markup));
  if (!match) throw new Error(`${label} is not a canonical self-closing <${qualifiedName}> leaf.`);
  const attributes = [];
  const seen = new Set();
  let rest = match[1];
  while (rest) {
    const attribute = /^\s+([:\w.-]+)="([^"]*)"/.exec(rest);
    if (!attribute) throw new Error(`${label} has unsupported XML attribute syntax.`);
    const [, name, value] = attribute;
    if (seen.has(name)) throw new Error(`${label} has a duplicate ${name} attribute.`);
    seen.add(name);
    attributes.push({ name, value });
    rest = rest.slice(attribute[0].length);
  }
  return { qualifiedName, attributes };
}

function replaceCanonicalDescription(markup, qualifiedName, replacement, label) {
  const parsed = parseCanonicalSelfClosingElement(markup, qualifiedName, label);
  let descriptions = 0;
  const attributes = parsed.attributes.map((attribute) => {
    if (attribute.name !== "descr") return attribute;
    descriptions += 1;
    return { name: attribute.name, value: encodeXmlAttribute(replacement) };
  });
  if (descriptions !== 1) throw new Error(`${label} must contain exactly one descr attribute.`);
  return `<${qualifiedName}${attributes.map(({ name, value }) => ` ${name}="${value}"`).join("")} />`;
}

function canonicalDescription(markup, qualifiedName, label) {
  const parsed = parseCanonicalSelfClosingElement(markup, qualifiedName, label);
  const description = parsed.attributes.filter((attribute) => attribute.name === "descr");
  if (description.length !== 1) throw new Error(`${label} must contain exactly one descr attribute.`);
  return decodeCanonicalXmlAttribute(description[0].value, `${label} descr`);
}

function exactlyOne(markups, label) {
  if (markups.length !== 1) throw new Error(`${label} requires exactly one matching element; found ${markups.length}.`);
  return markups[0];
}

function directBodyElements(xml, label) {
  const source = String(xml);
  const bodyMatch = /<w:body\b[^>]*>([\s\S]*)<\/w:body>/.exec(source);
  if (!bodyMatch) throw new Error(`${label} has no canonical w:body container.`);
  const bodyStart = (bodyMatch.index ?? 0) + bodyMatch[0].indexOf(">") + 1;
  const inner = bodyMatch[1];
  const elements = [];
  const stack = [];
  for (const match of inner.matchAll(/<\/?[\w:.-]+\b[^>]*>/g)) {
    const token = match[0];
    if (/^<\?/.test(token) || /^<!/.test(token)) throw new Error(`${label} has unsupported markup inside w:body.`);
    const closing = /^<\/([\w:.-]+)\s*>$/.exec(token);
    if (closing) {
      const current = stack.pop();
      if (!current || current.name !== closing[1]) throw new Error(`${label} has an unbalanced ${token} element.`);
      if (!stack.length) elements.push({ name: current.name, offset: bodyStart + current.offset, xml: inner.slice(current.offset, (match.index ?? 0) + token.length) });
      continue;
    }
    const opening = /^<([\w:.-]+)\b[^>]*>$/.exec(token);
    if (!opening) throw new Error(`${label} has unsupported XML token ${token}.`);
    const selfClosing = /\/>$/.test(token);
    if (selfClosing) {
      if (!stack.length) elements.push({ name: opening[1], offset: bodyStart + (match.index ?? 0), xml: token });
      continue;
    }
    stack.push({ name: opening[1], offset: match.index ?? 0 });
  }
  if (stack.length) throw new Error(`${label} has an unclosed ${stack.at(-1).name} element.`);
  return elements;
}

function selfClosingMatches(xml, qualifiedName) {
  return [...String(xml).matchAll(new RegExp(`<${qualifiedName}\\b[^>]*\\/>`, "g"))].map((match) => match[0]);
}

function compositeMatches(xml, qualifiedName) {
  return [...String(xml).matchAll(new RegExp(`<${qualifiedName}\\b[^>]*>[\\s\\S]*?<\\/${qualifiedName}>`, "g"))].map((match) => match[0]);
}

function rawImageParagraphs(xml, label) {
  const images = [];
  for (const element of directBodyElements(xml, label)) {
    if (element.name !== "w:p") continue;
    const drawingCount = compositeMatches(element.xml, "w:drawing").length;
    if (!drawingCount) continue;
    if (drawingCount !== 1) throw new Error(`${label} image paragraph requires exactly one w:drawing.`);
    const drawing = compositeMatches(element.xml, "w:drawing")[0];
    const inline = compositeMatches(drawing, "wp:inline");
    const anchor = compositeMatches(drawing, "wp:anchor");
    if (inline.length + anchor.length !== 1) throw new Error(`${label} image paragraph requires exactly one wp:inline or wp:anchor.`);
    const container = inline[0] || anchor[0];
    const docProperties = selfClosingMatches(container, "wp:docPr");
    const nonVisualProperties = selfClosingMatches(container, "pic:cNvPr");
    const docPrMarkup = exactlyOne(docProperties, `${label} wp:docPr`);
    const nonVisualMarkup = exactlyOne(nonVisualProperties, `${label} pic:cNvPr`);
    const docPrDescription = canonicalDescription(docPrMarkup, "wp:docPr", `${label} wp:docPr`);
    const nonVisualDescription = canonicalDescription(nonVisualMarkup, "pic:cNvPr", `${label} pic:cNvPr`);
    if (docPrDescription !== nonVisualDescription) {
      throw new Error(`${label} wp:docPr and pic:cNvPr descriptions do not agree.`);
    }
    images.push({
      offset: element.offset,
      paragraph: element.xml,
      docPrMarkup,
      nonVisualMarkup,
      alt: docPrDescription,
      placement: inline.length ? "inline" : "floating",
    });
  }
  return images;
}

function canonicalizeImageResidual(xml, label) {
  return String(xml).replace(/<[^>]+>/g, (tag) => {
    if (/^<\?/.test(tag) || /^<!/.test(tag) || /^<\//.test(tag)) return tag;
    const match = /^<([\w:.-]+)((?:\s+[:\w.-]+="[^"]*")*)\s*(\/?)>$/.exec(tag);
    if (!match) throw new Error(`${label} contains unsupported XML markup during residual comparison.`);
    const [, name, rawAttributes, slash] = match;
    const attributes = [];
    let rest = rawAttributes;
    while (rest) {
      const attribute = /^\s+([:\w.-]+)="([^"]*)"/.exec(rest);
      if (!attribute) throw new Error(`${label} contains unsupported XML attributes during residual comparison.`);
      const [, attributeName, value] = attribute;
      if (MOVABLE_NAMESPACES.has(attributeName)) {
        if (MOVABLE_NAMESPACES.get(attributeName) !== value) {
          throw new Error(`${label} changes the ${attributeName} namespace binding.`);
        }
      } else {
        attributes.push([attributeName, value]);
      }
      rest = rest.slice(attribute[0].length);
    }
    attributes.sort(([left], [right]) => left.localeCompare(right));
    const suffix = attributes.length ? ` ${attributes.map(([attributeName, value]) => `${attributeName}="${value}"`).join(" ")}` : "";
    return `<${name}${suffix}${slash}>`;
  });
}

function imageSnapshot(block, blockIndex, imageOrdinal) {
  if (block.kind !== "image" || typeof block.dataUrl !== "string" || !/^data:image\/(?:png|jpeg);base64,[A-Za-z0-9+/=]+$/i.test(block.dataUrl)) {
    throw new Error(`Document image block ${blockIndex} has no imported embedded PNG or JPEG data.`);
  }
  return {
    id: block.id,
    blockIndex,
    imageOrdinal,
    name: block.name || "",
    alt: String(block.alt || ""),
    styleId: block.styleId || "",
    widthPx: Number(block.widthPx),
    heightPx: Number(block.heightPx),
    placement: block.placement ? structuredClone(block.placement) : undefined,
    dataUrlSha256: sha256(Buffer.from(block.dataUrl, "utf8")),
  };
}

function imageProjection(document) {
  let imageOrdinal = 0;
  return document.blocks.flatMap((block, blockIndex) => {
    if (block.kind !== "image") return [];
    const snapshot = imageSnapshot(block, blockIndex, imageOrdinal);
    imageOrdinal += 1;
    return [snapshot];
  });
}

function selectImage(document, { imageBlockIndex, expectedAlt }) {
  const blockIndex = boundedIndex(imageBlockIndex, "imageBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block) throw new Error("imageBlockIndex is outside the imported document.");
  if (block.kind !== "image") throw new Error("imageBlockIndex does not identify an imported image block.");
  if (document.resolve(block.id) !== block) throw new Error("Selected image locator did not resolve to the inspected object.");
  const actualAlt = normalizeAltText(String(block.alt || ""), "selected image alt");
  if (actualAlt !== expectedAlt) {
    throw new Error(`Selected image alternative text does not match the expected source value: expected ${JSON.stringify(expectedAlt)}, observed ${JSON.stringify(actualAlt)}.`);
  }
  const imageOrdinal = document.blocks.slice(0, blockIndex + 1).filter((item) => item.kind === "image").length - 1;
  return { block, blockIndex, imageOrdinal, snapshot: imageSnapshot(block, blockIndex, imageOrdinal) };
}

function normalizedTargetImageXml(xml, selected, expectedImageCount, expectedAlt, label) {
  const images = rawImageParagraphs(xml, label);
  if (images.length !== expectedImageCount) {
    throw new Error(`${label} has ${images.length} canonical body images, but import exposed ${expectedImageCount}; refusing an ambiguous source-to-model mapping.`);
  }
  const target = images[selected.imageOrdinal];
  if (!target) throw new Error(`${label} has no canonical image at ordinal ${selected.imageOrdinal}.`);
  if (target.alt !== expectedAlt) {
    throw new Error(`${label} native alternative text does not match the expected source value: expected ${JSON.stringify(expectedAlt)}, observed ${JSON.stringify(target.alt)}.`);
  }
  const maskedParagraph = target.paragraph
    .replace(target.docPrMarkup, replaceCanonicalDescription(target.docPrMarkup, "wp:docPr", "officeKitAltTextMasked", `${label} wp:docPr`))
    .replace(target.nonVisualMarkup, replaceCanonicalDescription(target.nonVisualMarkup, "pic:cNvPr", "officeKitAltTextMasked", `${label} pic:cNvPr`));
  return {
    placement: target.placement,
    normalized: canonicalizeImageResidual(
      `${xml.slice(0, target.offset)}${maskedParagraph}${xml.slice(target.offset + target.paragraph.length)}`,
      label,
    ),
  };
}

function assertValidReplacement(document) {
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Replacement image alternative text fails document verification: ${verification.ndjson}`);
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes one recognized imported DOCX image's complete alternative-text value
 * through the public DocumentFile path. The target's source block index and
 * exact previous alternative text bind the operation; only its paired
 * wp:docPr/pic:cNvPr descr leaves in word/document.xml may differ.
 */
export async function editImportedImageAltText({
  inputPath,
  outputPath,
  auditPath,
  imageBlockIndex,
  expectedAlt,
  replacementAlt,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizeAltText(expectedAlt, "expectedAlt");
  const replacement = normalizeAltText(replacementAlt, "replacementAlt");
  if (expected === replacement) throw new Error("replacementAlt must differ from expectedAlt.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const sourceImages = imageProjection(document);
  const selected = selectImage(document, { imageBlockIndex, expectedAlt: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizedTargetImageXml(sourceXml, selected, sourceImages.length, expected, "source target image");
  if (sourceResidual.placement !== (selected.snapshot.placement ? "floating" : "inline")) {
    throw new Error("Source raw image placement does not match the imported image placement profile.");
  }
  selected.block.alt = replacement;
  assertValidReplacement(document);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound image alternative-text edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound image alternative-text edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }
    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizedTargetImageXml(outputXml, selected, sourceImages.length, replacement, "output target image");
    if (outputResidual.placement !== sourceResidual.placement || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Image alternative-text edit changed word/document.xml outside the bound paired description leaves.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectImage(reimported, { imageBlockIndex: selected.blockIndex, expectedAlt: replacement });
    const afterImages = imageProjection(reimported);
    const expectedImages = structuredClone(sourceImages);
    const expectedImage = expectedImages.find((image) => image.id === selected.snapshot.id);
    if (!expectedImage) throw new Error("Selected image disappeared from the imported image projection.");
    expectedImage.alt = replacement;
    if (!equalJson(afterImages, expectedImages)) {
      throw new Error("DOCX export changed imported image identity or semantics outside the requested alternative text.");
    }
    if (roundTrip.snapshot.id !== selected.snapshot.id) throw new Error("Second import did not preserve the selected image identity.");
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
        type: "source-bound-image-alt-text-edit",
        target: { id: selected.snapshot.id, blockIndex: selected.blockIndex, imageOrdinal: selected.imageOrdinal },
        sourceAlt: expected,
        replacementAlt: replacement,
        retained: {
          name: selected.snapshot.name,
          styleId: selected.snapshot.styleId,
          widthPx: selected.snapshot.widthPx,
          heightPx: selected.snapshot.heightPx,
          placement: selected.snapshot.placement,
          dataUrlSha256: selected.snapshot.dataUrlSha256,
        },
      },
      validation: {
        changedParts: changed,
        imageAltTextXmlResidual: {
          ok: true,
          imageOrdinal: selected.imageOrdinal,
          normalizedSha256: sha256(Buffer.from(sourceResidual.normalized, "utf8")),
        },
        reimport: { ok: true, imageId: roundTrip.snapshot.id, imageBlockIndex: roundTrip.blockIndex, imageOrdinal: roundTrip.imageOrdinal, alt: replacement },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes one image alternative-text value only. It preserves the image asset, dimensions, placement, and layout; inspect a native Word or LibreOffice render before delivery and review the replacement description for author-intent accuracy."],
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

export function parseImageAltTextEditCli(argv) {
  const [inputPath, outputPath, auditPath, imageBlockIndex, expectedAlt, replacementAlt] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    imageBlockIndex: boundedIndex(imageBlockIndex, "imageBlockIndex"),
    expectedAlt: normalizeAltText(expectedAlt, "expectedAlt"),
    replacementAlt: normalizeAltText(replacementAlt, "replacementAlt"),
  };
}

export function imageAltTextCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedImageAltText(parseImageAltTextEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(imageAltTextCliOutput(result)));
}
