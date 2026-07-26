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
import {
  boundedIndex,
  sectionProjection,
  sectionProperties,
  selectCanonicalSection,
  wordAttributes,
} from "../artifact_tool/_source_bound_sections.mjs";

const MARGIN_KEYS = ["top", "right", "bottom", "left", "gutter"];
const PAGE_MARGIN_KEYS = new Set([...MARGIN_KEYS, "header", "footer"]);
const MAX_TWIPS = 2_147_483_647;

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function normalizeMargins(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unsupported = Object.keys(value).filter((key) => !MARGIN_KEYS.includes(key));
  if (unsupported.length) throw new TypeError(`${label} has unsupported properties: ${unsupported.join(", ")}.`);
  const result = {};
  for (const key of MARGIN_KEYS) {
    if (!Object.hasOwn(value, key)) throw new TypeError(`${label}.${key} is required.`);
    const twips = Number(value[key]);
    if (!Number.isSafeInteger(twips) || twips < 0 || twips > MAX_TWIPS) {
      throw new TypeError(`${label}.${key} must be an integer from 0 through ${MAX_TWIPS}.`);
    }
    result[key] = twips;
  }
  return result;
}

function canonicalTwips(value, label) {
  if (!/^(?:0|[1-9]\d*)$/.test(String(value))) throw new Error(`${label} must be a non-negative canonical integer.`);
  const twips = Number(value);
  if (!Number.isSafeInteger(twips) || twips > MAX_TWIPS) throw new Error(`${label} is outside the supported twips range.`);
  return twips;
}

function canonicalPageMarginsLeaf(sectionXml, label) {
  const leaves = [...String(sectionXml).matchAll(/<w:pgMar\b[^>]*\/>/g)];
  if (leaves.length !== 1) throw new Error(`${label} must contain exactly one canonical w:pgMar leaf; found ${leaves.length}.`);
  const tag = leaves[0][0];
  const attributes = wordAttributes(tag, label);
  const unknown = Object.keys(attributes).filter((key) => !PAGE_MARGIN_KEYS.has(key));
  if (unknown.length) throw new Error(`${label} has unsupported w:pgMar attributes: ${unknown.join(", ")}.`);
  const values = {};
  for (const key of MARGIN_KEYS) {
    if (!Object.hasOwn(attributes, key)) throw new Error(`${label} is missing w:${key}.`);
    values[key] = canonicalTwips(attributes[key], `${label} w:${key}`);
  }
  for (const key of ["header", "footer"]) {
    if (!Object.hasOwn(attributes, key)) throw new Error(`${label} is missing w:${key}.`);
    values[key] = canonicalTwips(attributes[key], `${label} w:${key}`);
  }
  return { tag, margins: Object.fromEntries(MARGIN_KEYS.map((key) => [key, values[key]])), header: values.header, footer: values.footer };
}

function rawSectionMargins(xml, sectionOrdinal, label) {
  const sections = sectionProperties(xml);
  if (!sections[sectionOrdinal]) throw new Error(`${label} is outside the native w:sectPr sequence.`);
  return { sections, section: sections[sectionOrdinal], ...canonicalPageMarginsLeaf(sections[sectionOrdinal].xml, label) };
}

function normalizeTargetMarginsXml(xml, sectionOrdinal, label) {
  const raw = rawSectionMargins(xml, sectionOrdinal, label);
  const maskedTag = `<w:pgMar officeKitMarginsMasked="true" w:header="${raw.header}" w:footer="${raw.footer}"/>`;
  const normalizedSection = raw.section.xml.replace(raw.tag, maskedTag);
  return {
    margins: raw.margins,
    header: raw.header,
    footer: raw.footer,
    normalized: canonicalizeXmlForResidual(`${xml.slice(0, raw.section.offset)}${normalizedSection}${xml.slice(raw.section.offset + raw.section.xml.length)}`, label),
    sectionCount: raw.sections.length,
  };
}

function selectSection(document, { sectionBlockIndex, expectedMargins }) {
  const selected = selectCanonicalSection(document, sectionBlockIndex);
  const actual = normalizeMargins(selected.block.margins, "selected section margins");
  if (!equalJson(actual, expectedMargins)) {
    throw new Error(`Selected section margins do not match the expected source value: expected ${JSON.stringify(expectedMargins)}, observed ${JSON.stringify(actual)}.`);
  }
  return selected;
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes exactly one imported canonical w:pgMar leaf through the public
 * DocumentFile path. Header/footer page distances remain source-bound canaries.
 */
export async function editImportedSectionMargins({
  inputPath,
  outputPath,
  auditPath,
  sectionBlockIndex,
  expectedMargins,
  replacementMargins,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizeMargins(expectedMargins, "expectedMargins");
  const replacement = normalizeMargins(replacementMargins, "replacementMargins");
  if (equalJson(expected, replacement)) throw new Error("replacementMargins must differ from expectedMargins.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectSection(document, { sectionBlockIndex, expectedMargins: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizeTargetMarginsXml(sourceXml, selected.sectionOrdinal, "source target section");
  if (!equalJson(sourceResidual.margins, expected)) {
    throw new Error("The raw source w:pgMar does not match the inspected section margins.");
  }
  const beforeSections = sectionProjection(document);
  selected.block.margins = replacement;

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound section margin edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound section margin edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }

    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizeTargetMarginsXml(outputXml, selected.sectionOrdinal, "output target section");
    if (!equalJson(outputResidual.margins, replacement)) {
      throw new Error("Exported target w:pgMar does not match the requested margin replacement.");
    }
    if (sourceResidual.sectionCount !== outputResidual.sectionCount || sourceResidual.header !== outputResidual.header
      || sourceResidual.footer !== outputResidual.footer || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Section margin edit changed word/document.xml outside the requested canonical w:pgMar margin attributes.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectSection(reimported, { sectionBlockIndex: selected.blockIndex, expectedMargins: replacement });
    const afterSections = sectionProjection(reimported);
    const expectedSections = structuredClone(beforeSections);
    const expectedSection = expectedSections.find((section) => section.id === selected.snapshot.id);
    if (!expectedSection) throw new Error("Selected section disappeared from the imported section projection.");
    expectedSection.margins = replacement;
    if (!equalJson(afterSections, expectedSections)) {
      throw new Error("DOCX export changed imported section identity or settings outside the requested margin leaf.");
    }
    if (roundTrip.snapshot.id !== selected.snapshot.id || !roundTrip.block.editable) {
      throw new Error("Second import did not preserve the selected section identity or editable canonical boundary.");
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
        type: "source-bound-section-margin-edit",
        target: {
          id: selected.snapshot.id,
          blockIndex: selected.blockIndex,
          sectionOrdinal: selected.sectionOrdinal,
        },
        sourceMargins: expected,
        replacementMargins: replacement,
      },
      validation: {
        changedParts: changed,
        marginsXmlResidual: {
          ok: true,
          sectionOrdinal: selected.sectionOrdinal,
          normalizedSha256: sha256(Buffer.from(sourceResidual.normalized, "utf8")),
          headerTwips: sourceResidual.header,
          footerTwips: sourceResidual.footer,
        },
        reimport: {
          ok: true,
          sectionId: roundTrip.snapshot.id,
          sectionBlockIndex: roundTrip.blockIndex,
          sectionOrdinal: roundTrip.sectionOrdinal,
          editable: true,
          margins: replacement,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes only the selected section's canonical page margins. The model SVG is structural planning evidence; inspect a native Word or LibreOffice render before delivery."],
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

function parseJsonMargins(value, label) {
  try {
    return JSON.parse(requiredText(value, label));
  } catch (error) {
    throw new TypeError(`${label} must be JSON for a margins object: ${error.message}`);
  }
}

export function parseSectionMarginEditCli(argv) {
  const [inputPath, outputPath, auditPath, sectionBlockIndex, expectedMargins, replacementMargins] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    sectionBlockIndex: boundedIndex(sectionBlockIndex, "sectionBlockIndex"),
    expectedMargins: parseJsonMargins(expectedMargins, "expectedMargins"),
    replacementMargins: parseJsonMargins(replacementMargins, "replacementMargins"),
  };
}

export function sectionMarginCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedSectionMargins(parseSectionMarginEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(sectionMarginCliOutput(result)));
}
