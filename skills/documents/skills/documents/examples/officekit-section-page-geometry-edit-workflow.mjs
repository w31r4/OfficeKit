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

const PAGE_SIZE_KEYS = new Set(["w", "h", "orient"]);
const PAGE_GEOMETRY_KEYS = new Set(["orientation", "pageSize"]);
const MAX_TWIPS = 31_680;

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function canonicalPositiveTwips(value, label) {
  if (!/^[1-9]\d*$/.test(String(value))) throw new Error(`${label} must be a positive canonical integer.`);
  const twips = Number(value);
  if (!Number.isSafeInteger(twips) || twips > MAX_TWIPS) {
    throw new Error(`${label} must be an integer from 1 through ${MAX_TWIPS}.`);
  }
  return twips;
}

function normalizePageSize(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unsupported = Object.keys(value).filter((key) => key !== "widthTwips" && key !== "heightTwips");
  if (unsupported.length) throw new TypeError(`${label} has unsupported properties: ${unsupported.join(", ")}.`);
  const result = {};
  for (const key of ["widthTwips", "heightTwips"]) {
    if (!Object.hasOwn(value, key)) throw new TypeError(`${label}.${key} is required.`);
    const twips = Number(value[key]);
    if (!Number.isSafeInteger(twips) || twips < 1 || twips > MAX_TWIPS) {
      throw new TypeError(`${label}.${key} must be an integer from 1 through ${MAX_TWIPS}.`);
    }
    result[key] = twips;
  }
  return result;
}

function normalizePageGeometry(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unsupported = Object.keys(value).filter((key) => !PAGE_GEOMETRY_KEYS.has(key));
  if (unsupported.length) throw new TypeError(`${label} has unsupported properties: ${unsupported.join(", ")}.`);
  const orientation = String(value.orientation || "");
  if (!new Set(["portrait", "landscape"]).has(orientation)) {
    throw new TypeError(`${label}.orientation must be portrait or landscape.`);
  }
  return { orientation, pageSize: normalizePageSize(value.pageSize, `${label}.pageSize`) };
}

function canonicalPageSizeLeaf(sectionXml, label) {
  const leaves = [...String(sectionXml).matchAll(/<w:pgSz\b[^>]*\/>/g)];
  if (leaves.length !== 1) throw new Error(`${label} must contain exactly one canonical w:pgSz leaf; found ${leaves.length}.`);
  const tag = leaves[0][0];
  const attributes = wordAttributes(tag, label);
  const unknown = Object.keys(attributes).filter((key) => !PAGE_SIZE_KEYS.has(key));
  if (unknown.length) throw new Error(`${label} has unsupported w:pgSz attributes: ${unknown.join(", ")}.`);
  for (const key of PAGE_SIZE_KEYS) {
    if (!Object.hasOwn(attributes, key)) throw new Error(`${label} is missing w:${key}.`);
  }
  if (!new Set(["portrait", "landscape"]).has(attributes.orient)) {
    throw new Error(`${label} has an unsupported w:orient value.`);
  }
  return {
    tag,
    geometry: {
      orientation: attributes.orient,
      pageSize: {
        widthTwips: canonicalPositiveTwips(attributes.w, `${label} w:w`),
        heightTwips: canonicalPositiveTwips(attributes.h, `${label} w:h`),
      },
    },
  };
}

function rawSectionPageGeometry(xml, sectionOrdinal, label) {
  const sections = sectionProperties(xml);
  if (!sections[sectionOrdinal]) throw new Error(`${label} is outside the native w:sectPr sequence.`);
  return { sections, section: sections[sectionOrdinal], ...canonicalPageSizeLeaf(sections[sectionOrdinal].xml, label) };
}

function normalizeTargetPageGeometryXml(xml, sectionOrdinal, label) {
  const raw = rawSectionPageGeometry(xml, sectionOrdinal, label);
  const normalizedSection = raw.section.xml.replace(raw.tag, "<w:pgSz officeKitGeometryMasked=\"true\"/>");
  return {
    geometry: raw.geometry,
    normalized: canonicalizeXmlForResidual(`${xml.slice(0, raw.section.offset)}${normalizedSection}${xml.slice(raw.section.offset + raw.section.xml.length)}`, label),
    sectionCount: raw.sections.length,
  };
}

function selectSection(document, { sectionBlockIndex, expectedPageGeometry }) {
  const selected = selectCanonicalSection(document, sectionBlockIndex);
  const actual = normalizePageGeometry({
    orientation: selected.block.orientation,
    pageSize: selected.block.pageSize,
  }, "selected section page geometry");
  if (!equalJson(actual, expectedPageGeometry)) {
    throw new Error(`Selected section page geometry does not match the expected source value: expected ${JSON.stringify(expectedPageGeometry)}, observed ${JSON.stringify(actual)}.`);
  }
  return selected;
}

function assertValidReplacement(document) {
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) {
    throw new Error(`Replacement page geometry fails document verification: ${verification.ndjson}`);
  }
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes exactly one imported canonical w:pgSz leaf through the public
 * DocumentFile path. Orientation and both dimensions are one atomic geometry
 * value, never independently patched.
 */
export async function editImportedSectionPageGeometry({
  inputPath,
  outputPath,
  auditPath,
  sectionBlockIndex,
  expectedPageGeometry,
  replacementPageGeometry,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizePageGeometry(expectedPageGeometry, "expectedPageGeometry");
  const replacement = normalizePageGeometry(replacementPageGeometry, "replacementPageGeometry");
  if (equalJson(expected, replacement)) throw new Error("replacementPageGeometry must differ from expectedPageGeometry.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectSection(document, { sectionBlockIndex, expectedPageGeometry: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizeTargetPageGeometryXml(sourceXml, selected.sectionOrdinal, "source target section");
  if (!equalJson(sourceResidual.geometry, expected)) {
    throw new Error("The raw source w:pgSz does not match the inspected section page geometry.");
  }
  const beforeSections = sectionProjection(document);
  selected.block.orientation = replacement.orientation;
  selected.block.pageSize = structuredClone(replacement.pageSize);
  assertValidReplacement(document);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound section page-geometry edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound section page-geometry edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }

    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizeTargetPageGeometryXml(outputXml, selected.sectionOrdinal, "output target section");
    if (!equalJson(outputResidual.geometry, replacement)) {
      throw new Error("Exported target w:pgSz does not match the requested page-geometry replacement.");
    }
    if (sourceResidual.sectionCount !== outputResidual.sectionCount || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Section page-geometry edit changed word/document.xml outside the requested canonical w:pgSz leaf.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectSection(reimported, { sectionBlockIndex: selected.blockIndex, expectedPageGeometry: replacement });
    const afterSections = sectionProjection(reimported);
    const expectedSections = structuredClone(beforeSections);
    const expectedSection = expectedSections.find((section) => section.id === selected.snapshot.id);
    if (!expectedSection) throw new Error("Selected section disappeared from the imported section projection.");
    expectedSection.orientation = replacement.orientation;
    expectedSection.pageSize = structuredClone(replacement.pageSize);
    if (!equalJson(afterSections, expectedSections)) {
      throw new Error("DOCX export changed imported section identity or settings outside the requested page-size leaf.");
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
        type: "source-bound-section-page-geometry-edit",
        target: {
          id: selected.snapshot.id,
          blockIndex: selected.blockIndex,
          sectionOrdinal: selected.sectionOrdinal,
        },
        sourcePageGeometry: expected,
        replacementPageGeometry: replacement,
      },
      validation: {
        changedParts: changed,
        pageSizeXmlResidual: {
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
          pageGeometry: replacement,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes only the selected section's canonical page geometry. The model SVG is structural planning evidence; inspect a native Word or LibreOffice render before delivery."],
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

function parseJsonPageGeometry(value, label) {
  try {
    return JSON.parse(requiredText(value, label));
  } catch (error) {
    throw new TypeError(`${label} must be JSON for a page-geometry object: ${error.message}`);
  }
}

export function parseSectionPageGeometryEditCli(argv) {
  const [inputPath, outputPath, auditPath, sectionBlockIndex, expectedPageGeometry, replacementPageGeometry] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    sectionBlockIndex: boundedIndex(sectionBlockIndex, "sectionBlockIndex"),
    expectedPageGeometry: parseJsonPageGeometry(expectedPageGeometry, "expectedPageGeometry"),
    replacementPageGeometry: parseJsonPageGeometry(replacementPageGeometry, "replacementPageGeometry"),
  };
}

export function sectionPageGeometryCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedSectionPageGeometry(parseSectionPageGeometryEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(sectionPageGeometryCliOutput(result)));
}
