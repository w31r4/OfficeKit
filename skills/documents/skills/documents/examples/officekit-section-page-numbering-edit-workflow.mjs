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

const PAGE_NUMBER_FORMATS = new Set(["decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"]);

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

function normalizePageNumbering(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unsupported = Object.keys(value).filter((key) => key !== "start" && key !== "format");
  if (unsupported.length) throw new TypeError(`${label} has unsupported properties: ${unsupported.join(", ")}.`);
  const result = {};
  if (Object.hasOwn(value, "start")) {
    const start = Number(value.start);
    if (!Number.isInteger(start) || start < 0 || start > 2_147_483_647) {
      throw new TypeError(`${label}.start must be an integer from 0 through 2147483647.`);
    }
    result.start = start;
  }
  if (Object.hasOwn(value, "format")) {
    const format = String(value.format);
    if (!PAGE_NUMBER_FORMATS.has(format)) {
      throw new TypeError(`${label}.format must be decimal, upperRoman, lowerRoman, upperLetter, or lowerLetter.`);
    }
    result.format = format;
  }
  if (!Object.keys(result).length) throw new TypeError(`${label} must include start or format.`);
  return result;
}

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function xmlAttributes(opening = "") {
  const result = {};
  for (const match of String(opening).matchAll(/([:\w.-]+)="([^"]*)"/g)) {
    result[match[1].split(":").at(-1)] = match[2];
  }
  return result;
}

function sectionProperties(xml) {
  return [...String(xml).matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((match, index) => ({
    index,
    xml: match[0],
    offset: match.index,
  }));
}

function canonicalPageNumberingLeaf(sectionXml, label) {
  const leaves = [...String(sectionXml).matchAll(/<w:pgNumType\b[^>]*\/>/g)];
  if (leaves.length !== 1) throw new Error(`${label} must contain exactly one canonical w:pgNumType leaf; found ${leaves.length}.`);
  const tag = leaves[0][0];
  const attributes = xmlAttributes(tag);
  const unknown = Object.keys(attributes).filter((key) => key !== "start" && key !== "fmt");
  if (unknown.length) throw new Error(`${label} has unsupported w:pgNumType attributes: ${unknown.join(", ")}.`);
  const value = {};
  if (Object.hasOwn(attributes, "start")) {
    if (!/^(?:0|[1-9]\d*)$/.test(attributes.start)) throw new Error(`${label} has a non-canonical w:start value.`);
    const start = Number(attributes.start);
    if (!Number.isSafeInteger(start) || start > 2_147_483_647) throw new Error(`${label} has an out-of-range w:start value.`);
    value.start = start;
  }
  if (Object.hasOwn(attributes, "fmt")) {
    if (!PAGE_NUMBER_FORMATS.has(attributes.fmt)) throw new Error(`${label} has an unsupported w:fmt value.`);
    value.format = attributes.fmt;
  }
  if (!Object.keys(value).length) throw new Error(`${label} has an empty w:pgNumType leaf.`);
  return { tag, value };
}

function rawSectionPageNumbering(xml, sectionOrdinal, label) {
  const sections = sectionProperties(xml);
  if (!sections[sectionOrdinal]) throw new Error(`${label} is outside the native w:sectPr sequence.`);
  const leaf = canonicalPageNumberingLeaf(sections[sectionOrdinal].xml, label);
  return { sections, section: sections[sectionOrdinal], ...leaf };
}

function normalizeTargetPageNumberingXml(xml, sectionOrdinal, label) {
  const raw = rawSectionPageNumbering(xml, sectionOrdinal, label);
  const normalizedSection = raw.section.xml.replace(raw.tag, "<w:pgNumType officeKitMasked=\"true\"/>");
  return {
    value: raw.value,
    normalized: canonicalizeXmlForResidual(`${xml.slice(0, raw.section.offset)}${normalizedSection}${xml.slice(raw.section.offset + raw.section.xml.length)}`, label),
    sectionCount: raw.sections.length,
  };
}

function sectionSnapshot(block, blockIndex, sectionOrdinal) {
  return {
    id: block.id,
    blockIndex,
    sectionOrdinal,
    name: block.name || "",
    editable: block.editable,
    breakType: block.breakType,
    orientation: block.orientation,
    pageSize: block.pageSize,
    margins: block.margins,
    columns: block.columns,
    pageNumbering: block.pageNumbering,
    lineNumbering: block.lineNumbering,
  };
}

function sectionProjection(document) {
  let sectionOrdinal = 0;
  return document.blocks.flatMap((block, blockIndex) => {
    if (block.kind !== "section") return [];
    const snapshot = sectionSnapshot(block, blockIndex, sectionOrdinal);
    sectionOrdinal += 1;
    return [snapshot];
  });
}

function selectSection(document, { sectionBlockIndex, expectedPageNumbering }) {
  const blockIndex = boundedIndex(sectionBlockIndex, "sectionBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block) throw new Error("sectionBlockIndex is outside the imported document.");
  if (block.kind !== "section") throw new Error("sectionBlockIndex does not identify an imported section block.");
  if (!block.editable) throw new Error("The selected section is source-bound and read-only; its native w:sectPr graph is outside this workflow's canonical profile.");
  if (document.resolve(block.id) !== block) throw new Error("The selected section locator did not resolve to the inspected object.");
  const actual = normalizePageNumbering(block.pageNumbering, "selected section pageNumbering");
  if (!equalJson(actual, expectedPageNumbering)) {
    throw new Error(`Selected section pageNumbering does not match the expected source value: expected ${JSON.stringify(expectedPageNumbering)}, observed ${JSON.stringify(actual)}.`);
  }
  const sectionOrdinal = document.blocks.slice(0, blockIndex + 1).filter((item) => item.kind === "section").length - 1;
  return { block, blockIndex, sectionOrdinal, snapshot: sectionSnapshot(block, blockIndex, sectionOrdinal) };
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes one canonical, imported w:pgNumType leaf through the public
 * DocumentFile path. It deliberately does not add PAGE fields or refresh their
 * cached display values: those are pagination-host responsibilities.
 */
export async function editImportedSectionPageNumbering({
  inputPath,
  outputPath,
  auditPath,
  sectionBlockIndex,
  expectedPageNumbering,
  replacementPageNumbering,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizePageNumbering(expectedPageNumbering, "expectedPageNumbering");
  const replacement = normalizePageNumbering(replacementPageNumbering, "replacementPageNumbering");
  if (equalJson(expected, replacement)) throw new Error("replacementPageNumbering must differ from expectedPageNumbering.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectSection(document, { sectionBlockIndex, expectedPageNumbering: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizeTargetPageNumberingXml(sourceXml, selected.sectionOrdinal, "source target section");
  if (!equalJson(sourceResidual.value, expected)) {
    throw new Error("The raw source w:pgNumType does not match the inspected section pageNumbering.");
  }
  const beforeSections = sectionProjection(document);
  selected.block.pageNumbering = replacement;

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound section page-numbering edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound section page-numbering edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }

    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizeTargetPageNumberingXml(outputXml, selected.sectionOrdinal, "output target section");
    if (!equalJson(outputResidual.value, replacement)) {
      throw new Error("Exported target w:pgNumType does not match the requested page-numbering replacement.");
    }
    if (sourceResidual.sectionCount !== outputResidual.sectionCount || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Section page-numbering edit changed word/document.xml outside the one requested canonical w:pgNumType leaf.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectSection(reimported, { sectionBlockIndex: selected.blockIndex, expectedPageNumbering: replacement });
    const afterSections = sectionProjection(reimported);
    const expectedSections = structuredClone(beforeSections);
    const expectedSection = expectedSections.find((section) => section.id === selected.snapshot.id);
    if (!expectedSection) throw new Error("Selected section disappeared from the imported section projection.");
    expectedSection.pageNumbering = replacement;
    if (!equalJson(afterSections, expectedSections)) {
      throw new Error("DOCX export changed imported section identity or settings outside the requested page-numbering leaf.");
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
        type: "source-bound-section-page-numbering-edit",
        target: {
          id: selected.snapshot.id,
          blockIndex: selected.blockIndex,
          sectionOrdinal: selected.sectionOrdinal,
        },
        sourcePageNumbering: expected,
        replacementPageNumbering: replacement,
      },
      validation: {
        changedParts: changed,
        pageNumberingXmlResidual: {
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
          pageNumbering: replacement,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes only the selected section's canonical w:pgNumType metadata. It does not add PAGE fields or refresh cached page-number display text; inspect native Word or LibreOffice output before delivery."],
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

function parseJsonPageNumbering(value, label) {
  try {
    return JSON.parse(requiredText(value, label));
  } catch (error) {
    throw new TypeError(`${label} must be JSON for a pageNumbering object: ${error.message}`);
  }
}

export function parseSectionPageNumberingEditCli(argv) {
  const [inputPath, outputPath, auditPath, sectionBlockIndex, expectedPageNumbering, replacementPageNumbering] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    sectionBlockIndex: boundedIndex(sectionBlockIndex, "sectionBlockIndex"),
    expectedPageNumbering: parseJsonPageNumbering(expectedPageNumbering, "expectedPageNumbering"),
    replacementPageNumbering: parseJsonPageNumbering(replacementPageNumbering, "replacementPageNumbering"),
  };
}

export function sectionPageNumberingCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedSectionPageNumbering(parseSectionPageNumberingEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(sectionPageNumberingCliOutput(result)));
}
