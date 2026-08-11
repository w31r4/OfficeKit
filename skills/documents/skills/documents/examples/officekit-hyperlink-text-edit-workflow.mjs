import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { DocumentFile, FileBlob } from "office-kit";
import {
  DOCX_MIME,
  assertAbsent,
  canonicalizeXmlForResidual,
  changedParts,
  directBodyElements,
  packageVersion,
  publishNoReplace,
  readPackagePartText,
  requiredText,
  sha256,
} from "../artifact_tool/_source_bound_docx.mjs";

const MAX_LINK_TEXT_CHARS = 32_767;
const MAX_LINK_TARGET_CHARS = 4_096;

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

function normalizeLinkText(value, label, { allowEmpty = false } = {}) {
  if (typeof value !== "string" || value.length > MAX_LINK_TEXT_CHARS || /[\u0000-\u001f\u007f]/.test(value)) {
    throw new TypeError(`${label} must be a string of at most ${MAX_LINK_TEXT_CHARS} characters without controls.`);
  }
  if (!allowEmpty && !value.trim()) throw new TypeError(`${label} must contain visible non-whitespace text.`);
  return value;
}

function normalizeLinkTarget(value, label) {
  const target = requiredText(value, label);
  if (target.startsWith("#")) {
    const anchor = target.slice(1);
    if (!anchor || anchor.length > 255 || /[\u0000-\u001f\u007f]/.test(anchor)) {
      throw new TypeError(`${label} internal anchor must contain 1 through 255 characters without controls.`);
    }
    return `#${anchor}`;
  }
  let parsed;
  try { parsed = new URL(target); } catch { parsed = undefined; }
  if (!parsed || !new Set(["http:", "https:"]).has(parsed.protocol) || target.length > MAX_LINK_TARGET_CHARS || /[\u0000-\u001f\u007f]/.test(target)) {
    throw new TypeError(`${label} must be an absolute http(s) URI of at most ${MAX_LINK_TARGET_CHARS} characters or a #bookmark anchor.`);
  }
  return target;
}

function decodeXmlText(value, label) {
  const source = String(value);
  if (/<|&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-f]+);)/iu.test(source)) {
    throw new Error(`${label} contains unsupported XML text markup or escaping.`);
  }
  return source.replace(/&(amp|lt|gt|quot|apos|#(\d+)|#x([0-9a-f]+));/giu, (_match, entity, decimal, hexadecimal) => {
    if (decimal !== undefined || hexadecimal !== undefined) {
      const codePoint = Number.parseInt(decimal ?? hexadecimal, hexadecimal === undefined ? 10 : 16);
      if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff || codePoint >= 0xd800 && codePoint <= 0xdfff) {
        throw new Error(`${label} contains an invalid numeric XML entity.`);
      }
      return String.fromCodePoint(codePoint);
    }
    return ({ amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'" })[entity.toLowerCase()];
  });
}

function textLeafProfile(markup, label) {
  const paired = /^<w:t\b([^>]*)>([\s\S]*)<\/w:t>$/.exec(markup);
  const empty = paired ? undefined : /^<w:t\b([^>]*)\s*\/>$/.exec(markup);
  if (!paired && !empty) throw new Error(`${label} is not one canonical w:t leaf.`);
  const rawAttributes = (paired || empty)[1];
  const attributes = [];
  const seen = new Set();
  let rest = rawAttributes.trim();
  while (rest) {
    const attribute = /^([:\w.-]+)="([^"]*)"\s*/.exec(rest);
    if (!attribute) throw new Error(`${label} has unsupported XML attribute syntax.`);
    const [, name, value] = attribute;
    if (seen.has(name)) throw new Error(`${label} has a duplicate ${name} attribute.`);
    seen.add(name);
    if (name !== "xml:space") attributes.push([name, value]);
    rest = rest.slice(attribute[0].length);
  }
  const suffix = attributes.length ? ` ${attributes.map(([name, value]) => `${name}="${value}"`).join(" ")}` : "";
  return {
    text: decodeXmlText(paired?.[2] ?? "", label),
    masked: `<w:t${suffix}>officeKitHyperlinkTextMasked</w:t>`,
  };
}

function rawHyperlinkParagraphs(xml, label) {
  const hyperlinks = [];
  for (const element of directBodyElements(xml, label)) {
    if (element.name !== "w:p" || !/<w:hyperlink\b/u.test(element.xml)) continue;
    const matches = [...element.xml.matchAll(/<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>/g)];
    if (matches.length !== 1) throw new Error(`${label} direct hyperlink paragraph requires exactly one non-empty w:hyperlink element.`);
    const hyperlink = matches[0];
    const textMatches = [...hyperlink[0].matchAll(/<w:t\b[^>]*(?:\/>|>[\s\S]*?<\/w:t>)/g)];
    if (textMatches.length !== 1) throw new Error(`${label} direct hyperlink paragraph requires exactly one w:t leaf.`);
    const textMatch = textMatches[0];
    const profile = textLeafProfile(textMatch[0], `${label} hyperlink text`);
    hyperlinks.push({
      offset: element.offset,
      paragraph: element.xml,
      textOffset: (hyperlink.index ?? 0) + (textMatch.index ?? 0),
      textMarkup: textMatch[0],
      text: profile.text,
      maskedTextMarkup: profile.masked,
    });
  }
  return hyperlinks;
}

function hyperlinkTarget(block) {
  return block.anchor ? `#${block.anchor}` : String(block.url || "");
}

function hyperlinkSnapshot(block, blockIndex, hyperlinkOrdinal) {
  return {
    id: block.id,
    blockIndex,
    hyperlinkOrdinal,
    name: block.name || "",
    styleId: block.styleId || "",
    relationshipId: block.relationshipId || "",
    text: String(block.text || ""),
    target: hyperlinkTarget(block),
    tooltip: block.tooltip ?? null,
    history: block.history !== false,
  };
}

function hyperlinkProjection(document) {
  let hyperlinkOrdinal = 0;
  return document.blocks.flatMap((block, blockIndex) => {
    if (block.kind !== "hyperlink") return [];
    const snapshot = hyperlinkSnapshot(block, blockIndex, hyperlinkOrdinal);
    hyperlinkOrdinal += 1;
    return [snapshot];
  });
}

function selectHyperlink(document, { hyperlinkBlockIndex, expectedTarget, expectedText }) {
  const blockIndex = boundedIndex(hyperlinkBlockIndex, "hyperlinkBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block || block.kind !== "hyperlink") throw new Error("hyperlinkBlockIndex does not identify an imported hyperlink block.");
  if (document.resolve(block.id) !== block) throw new Error("Selected hyperlink locator did not resolve to the inspected object.");
  const actualTarget = hyperlinkTarget(block);
  if (actualTarget !== expectedTarget) {
    throw new Error(`Selected hyperlink target does not match the expected source value: expected ${JSON.stringify(expectedTarget)}, observed ${JSON.stringify(actualTarget)}.`);
  }
  const actualText = String(block.text || "");
  if (actualText !== expectedText) {
    throw new Error(`Selected hyperlink text does not match the expected source value: expected ${JSON.stringify(expectedText)}, observed ${JSON.stringify(actualText)}.`);
  }
  const hyperlinkOrdinal = document.blocks.slice(0, blockIndex + 1).filter((item) => item.kind === "hyperlink").length - 1;
  return { block, blockIndex, hyperlinkOrdinal, snapshot: hyperlinkSnapshot(block, blockIndex, hyperlinkOrdinal) };
}

function normalizedTargetHyperlinkXml(xml, selected, expectedHyperlinkCount, expectedText, label) {
  const hyperlinks = rawHyperlinkParagraphs(xml, label);
  if (hyperlinks.length !== expectedHyperlinkCount) {
    throw new Error(`${label} has ${hyperlinks.length} direct native hyperlinks, but import exposed ${expectedHyperlinkCount}; refusing an ambiguous source-to-model mapping.`);
  }
  const target = hyperlinks[selected.hyperlinkOrdinal];
  if (!target) throw new Error(`${label} has no direct hyperlink at ordinal ${selected.hyperlinkOrdinal}.`);
  if (target.text !== expectedText) {
    throw new Error(`${label} native hyperlink text does not match the expected source value: expected ${JSON.stringify(expectedText)}, observed ${JSON.stringify(target.text)}.`);
  }
  const maskedParagraph = `${target.paragraph.slice(0, target.textOffset)}${target.maskedTextMarkup}${target.paragraph.slice(target.textOffset + target.textMarkup.length)}`;
  return canonicalizeXmlForResidual(
    `${xml.slice(0, target.offset)}${maskedParagraph}${xml.slice(target.offset + target.paragraph.length)}`,
    label,
  );
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/iu.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length, sha256: sha256(Buffer.from(svg, "utf8")) };
}

/**
 * Replaces the visible text of one recognized imported whole-paragraph DOCX
 * hyperlink. The block index, complete old text, and target bind the edit;
 * the target, formatting, relationship identity, and all non-target package
 * state remain source-owned.
 */
export async function editImportedHyperlinkText({
  inputPath,
  outputPath,
  auditPath,
  hyperlinkBlockIndex,
  expectedTarget,
  expectedText,
  replacementText,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const target = normalizeLinkTarget(expectedTarget, "expectedTarget");
  const expected = normalizeLinkText(expectedText, "expectedText", { allowEmpty: true });
  const replacement = normalizeLinkText(replacementText, "replacementText");
  if (expected === replacement) throw new Error("replacementText must differ from expectedText.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const sourceHyperlinks = hyperlinkProjection(document);
  const selected = selectHyperlink(document, {
    hyperlinkBlockIndex,
    expectedTarget: target,
    expectedText: expected,
  });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizedTargetHyperlinkXml(sourceXml, selected, sourceHyperlinks.length, expected, "source target hyperlink");
  selected.block.text = replacement;
  const sourceVerification = document.verify({ visualQa: true });
  if (!sourceVerification.ok) throw new Error(`Replacement hyperlink text fails document verification: ${sourceVerification.ndjson}`);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound hyperlink text edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound hyperlink text edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }
    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizedTargetHyperlinkXml(outputXml, selected, sourceHyperlinks.length, replacement, "output target hyperlink");
    if (outputResidual !== sourceResidual) {
      throw new Error("Hyperlink text edit changed word/document.xml outside the bound w:t text leaf and xml:space state.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectHyperlink(reimported, {
      hyperlinkBlockIndex: selected.blockIndex,
      expectedTarget: target,
      expectedText: replacement,
    });
    const afterHyperlinks = hyperlinkProjection(reimported);
    const expectedHyperlinks = structuredClone(sourceHyperlinks);
    const expectedHyperlink = expectedHyperlinks.find((item) => item.id === selected.snapshot.id);
    if (!expectedHyperlink) throw new Error("Selected hyperlink disappeared from the imported hyperlink projection.");
    expectedHyperlink.text = replacement;
    if (!equalJson(afterHyperlinks, expectedHyperlinks)) {
      throw new Error("DOCX export changed hyperlink identity or semantics outside the requested visible text.");
    }
    if (roundTrip.snapshot.id !== selected.snapshot.id) throw new Error("Second import did not preserve the selected hyperlink identity.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
    const accessibility = reimported.auditAccessibility({ maxChars: 200_000 });
    if (accessibility.issues.some((entry) => entry.id === selected.snapshot.id && entry.type === "hyperlinkTextMissing")) {
      throw new Error("Second import still reports the selected hyperlink as missing visible text.");
    }
    const render = await modelRender(reimported);
    const audit = {
      schema: "office-kit.docx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", noReplace: true },
      operation: {
        type: "source-bound-hyperlink-text-edit",
        target: {
          id: selected.snapshot.id,
          blockIndex: selected.blockIndex,
          hyperlinkOrdinal: selected.hyperlinkOrdinal,
          destination: target,
        },
        expectedText: expected,
        replacementText: replacement,
        retained: {
          name: selected.snapshot.name,
          styleId: selected.snapshot.styleId,
          relationshipId: selected.snapshot.relationshipId,
          tooltip: selected.snapshot.tooltip,
          history: selected.snapshot.history,
        },
      },
      validation: {
        changedParts: changed,
        hyperlinkTextXmlResidual: {
          ok: true,
          hyperlinkOrdinal: selected.hyperlinkOrdinal,
          normalizedSha256: sha256(Buffer.from(sourceResidual, "utf8")),
        },
        reimport: { ok: true, hyperlinkId: roundTrip.snapshot.id, blockIndex: roundTrip.blockIndex, text: replacement, target },
        accessibility: {
          selectedHyperlinkTextPresent: true,
          conformanceClaimed: false,
          documentMachineCheckPassed: accessibility.machineCheckPassed,
          manualReviewRequired: accessibility.manualReviewRequired,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["The replacement text still requires author review for destination purpose and surrounding context. This transaction does not change the destination, paragraph/run formatting, relationship graph, or claim Word/WCAG conformance."],
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

export function parseHyperlinkTextEditCli(argv) {
  const [inputPath, outputPath, auditPath, hyperlinkBlockIndex, expectedTarget, expectedText = "", replacementText] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    hyperlinkBlockIndex: boundedIndex(hyperlinkBlockIndex, "hyperlinkBlockIndex"),
    expectedTarget: normalizeLinkTarget(expectedTarget, "expectedTarget"),
    expectedText: normalizeLinkText(expectedText, "expectedText", { allowEmpty: true }),
    replacementText: normalizeLinkText(replacementText, "replacementText"),
  };
}

export function hyperlinkTextCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedHyperlinkText(parseHyperlinkTextEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(hyperlinkTextCliOutput(result)));
}
