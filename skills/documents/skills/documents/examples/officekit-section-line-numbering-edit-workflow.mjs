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

const LINE_NUMBER_RESTARTS = new Set(["newPage", "newSection", "continuous"]);
const LINE_NUMBER_KEYS = new Set(["countBy", "start", "distance", "restart"]);
const MAX_COUNT_BY = 32_767;
const MAX_START = 32_767;
const MAX_DISTANCE = 31_680;

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function normalizeLineNumbering(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unsupported = Object.keys(value).filter((key) => !LINE_NUMBER_KEYS.has(key));
  if (unsupported.length) throw new TypeError(`${label} has unsupported properties: ${unsupported.join(", ")}.`);
  if (!Object.hasOwn(value, "countBy")) throw new TypeError(`${label}.countBy is required.`);
  const countBy = Number(value.countBy);
  if (!Number.isSafeInteger(countBy) || countBy < 1 || countBy > MAX_COUNT_BY) {
    throw new TypeError(`${label}.countBy must be an integer from 1 through ${MAX_COUNT_BY}.`);
  }
  const result = { countBy };
  if (Object.hasOwn(value, "start")) {
    const start = Number(value.start);
    if (!Number.isSafeInteger(start) || start < 0 || start > MAX_START) {
      throw new TypeError(`${label}.start must be an integer from 0 through ${MAX_START}.`);
    }
    result.start = start;
  }
  if (Object.hasOwn(value, "distance")) {
    const distance = Number(value.distance);
    if (!Number.isSafeInteger(distance) || distance < 0 || distance > MAX_DISTANCE) {
      throw new TypeError(`${label}.distance must be an integer from 0 through ${MAX_DISTANCE}.`);
    }
    result.distance = distance;
  }
  if (Object.hasOwn(value, "restart")) {
    const restart = String(value.restart);
    if (!LINE_NUMBER_RESTARTS.has(restart)) {
      throw new TypeError(`${label}.restart must be newPage, newSection, or continuous.`);
    }
    result.restart = restart;
  }
  return result;
}

function canonicalUnsignedInteger(value, label, maximum) {
  if (!/^(?:0|[1-9]\d*)$/.test(String(value))) throw new Error(`${label} must be a canonical unsigned integer.`);
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > maximum) {
    throw new Error(`${label} must be an integer from 0 through ${maximum}.`);
  }
  return parsed;
}

function canonicalLineNumberingLeaf(sectionXml, label) {
  const leaves = [...String(sectionXml).matchAll(/<w:lnNumType\b[^>]*\/>/g)];
  if (leaves.length !== 1) throw new Error(`${label} must contain exactly one canonical w:lnNumType leaf; found ${leaves.length}.`);
  const tag = leaves[0][0];
  const attributes = wordAttributes(tag, label);
  const unknown = Object.keys(attributes).filter((key) => !LINE_NUMBER_KEYS.has(key));
  if (unknown.length) throw new Error(`${label} has unsupported w:lnNumType attributes: ${unknown.join(", ")}.`);
  const value = {
    countBy: Object.hasOwn(attributes, "countBy")
      ? canonicalUnsignedInteger(attributes.countBy, `${label} w:countBy`, MAX_COUNT_BY)
      : 1,
  };
  if (value.countBy < 1) throw new Error(`${label} w:countBy must be at least 1.`);
  if (Object.hasOwn(attributes, "start")) value.start = canonicalUnsignedInteger(attributes.start, `${label} w:start`, MAX_START);
  if (Object.hasOwn(attributes, "distance")) value.distance = canonicalUnsignedInteger(attributes.distance, `${label} w:distance`, MAX_DISTANCE);
  if (Object.hasOwn(attributes, "restart")) {
    if (!LINE_NUMBER_RESTARTS.has(attributes.restart)) {
      throw new Error(`${label} has an unsupported w:restart value.`);
    }
    value.restart = attributes.restart;
  }
  return { tag, value };
}

function rawSectionLineNumbering(xml, sectionOrdinal, label) {
  const sections = sectionProperties(xml);
  if (!sections[sectionOrdinal]) throw new Error(`${label} is outside the native w:sectPr sequence.`);
  return { sections, section: sections[sectionOrdinal], ...canonicalLineNumberingLeaf(sections[sectionOrdinal].xml, label) };
}

function normalizeTargetLineNumberingXml(xml, sectionOrdinal, label) {
  const raw = rawSectionLineNumbering(xml, sectionOrdinal, label);
  const normalizedSection = raw.section.xml.replace(raw.tag, "<w:lnNumType officeKitLineNumberingMasked=\"true\"/>");
  return {
    value: raw.value,
    normalized: canonicalizeXmlForResidual(`${xml.slice(0, raw.section.offset)}${normalizedSection}${xml.slice(raw.section.offset + raw.section.xml.length)}`, label),
    sectionCount: raw.sections.length,
  };
}

function selectSection(document, { sectionBlockIndex, expectedLineNumbering }) {
  const selected = selectCanonicalSection(document, sectionBlockIndex);
  const actual = normalizeLineNumbering(selected.block.lineNumbering, "selected section lineNumbering");
  if (!equalJson(actual, expectedLineNumbering)) {
    throw new Error(`Selected section lineNumbering does not match the expected source value: expected ${JSON.stringify(expectedLineNumbering)}, observed ${JSON.stringify(actual)}.`);
  }
  return selected;
}

function assertValidReplacement(document) {
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) {
    throw new Error(`Replacement line numbering fails document verification: ${verification.ndjson}`);
  }
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes one imported canonical w:lnNumType leaf through the public
 * DocumentFile path. It changes display cadence/offset/distance/restart
 * metadata only; a native pagination host remains authoritative for display.
 */
export async function editImportedSectionLineNumbering({
  inputPath,
  outputPath,
  auditPath,
  sectionBlockIndex,
  expectedLineNumbering,
  replacementLineNumbering,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizeLineNumbering(expectedLineNumbering, "expectedLineNumbering");
  const replacement = normalizeLineNumbering(replacementLineNumbering, "replacementLineNumbering");
  if (equalJson(expected, replacement)) throw new Error("replacementLineNumbering must differ from expectedLineNumbering.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectSection(document, { sectionBlockIndex, expectedLineNumbering: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizeTargetLineNumberingXml(sourceXml, selected.sectionOrdinal, "source target section");
  if (!equalJson(sourceResidual.value, expected)) {
    throw new Error("The raw source w:lnNumType does not match the inspected section lineNumbering.");
  }
  const beforeSections = sectionProjection(document);
  selected.block.lineNumbering = structuredClone(replacement);
  assertValidReplacement(document);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound section line-numbering edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound section line-numbering edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }

    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizeTargetLineNumberingXml(outputXml, selected.sectionOrdinal, "output target section");
    if (!equalJson(outputResidual.value, replacement)) {
      throw new Error("Exported target w:lnNumType does not match the requested line-numbering replacement.");
    }
    if (sourceResidual.sectionCount !== outputResidual.sectionCount || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Section line-numbering edit changed word/document.xml outside the one requested canonical w:lnNumType leaf.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectSection(reimported, { sectionBlockIndex: selected.blockIndex, expectedLineNumbering: replacement });
    const afterSections = sectionProjection(reimported);
    const expectedSections = structuredClone(beforeSections);
    const expectedSection = expectedSections.find((section) => section.id === selected.snapshot.id);
    if (!expectedSection) throw new Error("Selected section disappeared from the imported section projection.");
    expectedSection.lineNumbering = structuredClone(replacement);
    if (!equalJson(afterSections, expectedSections)) {
      throw new Error("DOCX export changed imported section identity or settings outside the requested line-numbering leaf.");
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
        type: "source-bound-section-line-numbering-edit",
        target: {
          id: selected.snapshot.id,
          blockIndex: selected.blockIndex,
          sectionOrdinal: selected.sectionOrdinal,
        },
        sourceLineNumbering: expected,
        replacementLineNumbering: replacement,
      },
      validation: {
        changedParts: changed,
        lineNumberingXmlResidual: {
          ok: true,
          sectionOrdinal: selected.sectionOrdinal,
          normalizedSha256: sha256(Buffer.from(sourceResidual.normalized, "utf8")),
        },
        reimport: {
          ok: true,
          sectionId: roundTrip.snapshot.id,
          sectionBlockIndex: roundTrip.blockIndex,
          sectionOrdinal: roundTrip.sectionOrdinal,
          editable: true,
          lineNumbering: replacement,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes only the selected section's canonical w:lnNumType metadata. It does not calculate or visually infer line-number display; inspect a native Word or LibreOffice render before delivery."],
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

function parseJsonLineNumbering(value, label) {
  try {
    return JSON.parse(requiredText(value, label));
  } catch (error) {
    throw new TypeError(`${label} must be JSON for a lineNumbering object: ${error.message}`);
  }
}

export function parseSectionLineNumberingEditCli(argv) {
  const [inputPath, outputPath, auditPath, sectionBlockIndex, expectedLineNumbering, replacementLineNumbering] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    sectionBlockIndex: boundedIndex(sectionBlockIndex, "sectionBlockIndex"),
    expectedLineNumbering: parseJsonLineNumbering(expectedLineNumbering, "expectedLineNumbering"),
    replacementLineNumbering: parseJsonLineNumbering(replacementLineNumbering, "replacementLineNumbering"),
  };
}

export function sectionLineNumberingCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedSectionLineNumbering(parseSectionLineNumberingEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(sectionLineNumberingCliOutput(result)));
}
