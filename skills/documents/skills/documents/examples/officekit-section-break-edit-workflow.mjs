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

const BREAK_TYPES = new Set(["nextPage", "continuous", "evenPage", "oddPage"]);

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function normalizeBreakType(value, label) {
  if (typeof value !== "string" || !BREAK_TYPES.has(value)) {
    throw new TypeError(`${label} must be nextPage, continuous, evenPage, or oddPage.`);
  }
  return value;
}

function canonicalSectionTypeLeaf(sectionXml, label) {
  const leaves = [...String(sectionXml).matchAll(/<w:type\b[^>]*\/>/g)];
  if (leaves.length !== 1) throw new Error(`${label} must contain exactly one canonical w:type leaf; found ${leaves.length}.`);
  const tag = leaves[0][0];
  const attributes = wordAttributes(tag, label);
  const unknown = Object.keys(attributes).filter((key) => key !== "val");
  if (unknown.length) throw new Error(`${label} has unsupported w:type attributes: ${unknown.join(", ")}.`);
  if (!Object.hasOwn(attributes, "val")) throw new Error(`${label} w:type is missing w:val.`);
  return { tag, value: normalizeBreakType(attributes.val, `${label} w:val`) };
}

function rawSectionBreakType(xml, sectionOrdinal, label) {
  const sections = sectionProperties(xml);
  if (!sections[sectionOrdinal]) throw new Error(`${label} is outside the native w:sectPr sequence.`);
  return { sections, section: sections[sectionOrdinal], ...canonicalSectionTypeLeaf(sections[sectionOrdinal].xml, label) };
}

function normalizeTargetSectionTypeXml(xml, sectionOrdinal, label) {
  const raw = rawSectionBreakType(xml, sectionOrdinal, label);
  const normalizedSection = raw.section.xml.replace(raw.tag, "<w:type officeKitSectionBreakMasked=\"true\"/>");
  return {
    value: raw.value,
    normalized: canonicalizeXmlForResidual(`${xml.slice(0, raw.section.offset)}${normalizedSection}${xml.slice(raw.section.offset + raw.section.xml.length)}`, label),
    sectionCount: raw.sections.length,
  };
}

function selectSection(document, { sectionBlockIndex, expectedBreakType }) {
  const selected = selectCanonicalSection(document, sectionBlockIndex);
  const actual = normalizeBreakType(selected.block.breakType, "selected section breakType");
  if (actual !== expectedBreakType) {
    throw new Error(`Selected section breakType does not match the expected source value: expected ${expectedBreakType}, observed ${actual}.`);
  }
  return selected;
}

function assertValidReplacement(document) {
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Replacement section break type fails document verification: ${verification.ndjson}`);
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes one imported canonical w:type leaf through the public DocumentFile
 * path. It does not add or remove a section break, move the boundary, or
 * calculate the pages a Word-compatible host will produce.
 */
export async function editImportedSectionBreakType({
  inputPath,
  outputPath,
  auditPath,
  sectionBlockIndex,
  expectedBreakType,
  replacementBreakType,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizeBreakType(expectedBreakType, "expectedBreakType");
  const replacement = normalizeBreakType(replacementBreakType, "replacementBreakType");
  if (expected === replacement) throw new Error("replacementBreakType must differ from expectedBreakType.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectSection(document, { sectionBlockIndex, expectedBreakType: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizeTargetSectionTypeXml(sourceXml, selected.sectionOrdinal, "source target section");
  if (sourceResidual.value !== expected) {
    throw new Error("The raw source w:type does not match the inspected section break type.");
  }
  const beforeSections = sectionProjection(document);
  selected.block.breakType = replacement;
  assertValidReplacement(document);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound section break-type edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound section break-type edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }

    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizeTargetSectionTypeXml(outputXml, selected.sectionOrdinal, "output target section");
    if (outputResidual.value !== replacement) {
      throw new Error("Exported target w:type does not match the requested section break-type replacement.");
    }
    if (sourceResidual.sectionCount !== outputResidual.sectionCount || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Section break-type edit changed word/document.xml outside the one requested canonical w:type leaf.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectSection(reimported, { sectionBlockIndex: selected.blockIndex, expectedBreakType: replacement });
    const afterSections = sectionProjection(reimported);
    const expectedSections = structuredClone(beforeSections);
    const expectedSection = expectedSections.find((section) => section.id === selected.snapshot.id);
    if (!expectedSection) throw new Error("Selected section disappeared from the imported section projection.");
    expectedSection.breakType = replacement;
    if (!equalJson(afterSections, expectedSections)) {
      throw new Error("DOCX export changed imported section identity or settings outside the requested section-break leaf.");
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
        type: "source-bound-section-break-type-edit",
        target: {
          id: selected.snapshot.id,
          blockIndex: selected.blockIndex,
          sectionOrdinal: selected.sectionOrdinal,
        },
        sourceBreakType: expected,
        replacementBreakType: replacement,
      },
      validation: {
        changedParts: changed,
        sectionTypeXmlResidual: {
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
          breakType: replacement,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes only the selected section's canonical w:type metadata. It does not move the section boundary or calculate or visually infer pagination; inspect a native Word or LibreOffice render before delivery."],
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

export function parseSectionBreakTypeEditCli(argv) {
  const [inputPath, outputPath, auditPath, sectionBlockIndex, expectedBreakType, replacementBreakType] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    sectionBlockIndex: boundedIndex(sectionBlockIndex, "sectionBlockIndex"),
    expectedBreakType,
    replacementBreakType,
  };
}

export function sectionBreakTypeCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedSectionBreakType(parseSectionBreakTypeEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(sectionBreakTypeCliOutput(result)));
}
