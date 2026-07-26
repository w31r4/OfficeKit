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

const COLUMN_KEYS = new Set(["count", "spacing", "separator", "definitions"]);
const COLUMN_DEFINITION_KEYS = new Set(["width", "spacing"]);
const EQUAL_WIDTH_ATTRIBUTES = new Set(["equalWidth", "num", "space", "sep"]);
const CUSTOM_WIDTH_ATTRIBUTES = new Set(["equalWidth", "sep"]);
const CUSTOM_COLUMN_ATTRIBUTES = new Set(["w", "space"]);
const MAX_COLUMNS = 45;
const MAX_TWIPS = 31_680;

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function normalizedPositiveInteger(value, label, maximum) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 1 || parsed > maximum) {
    throw new TypeError(`${label} must be an integer from 1 through ${maximum}.`);
  }
  return parsed;
}

function normalizedNonnegativeInteger(value, label, maximum) {
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed < 0 || parsed > maximum) {
    throw new TypeError(`${label} must be an integer from 0 through ${maximum}.`);
  }
  return parsed;
}

function normalizeColumns(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const unsupported = Object.keys(value).filter((key) => !COLUMN_KEYS.has(key));
  if (unsupported.length) throw new TypeError(`${label} has unsupported properties: ${unsupported.join(", ")}.`);
  if (Object.hasOwn(value, "definitions")) {
    if (Object.hasOwn(value, "count") || Object.hasOwn(value, "spacing")) {
      throw new TypeError(`${label} custom-width columns cannot combine definitions with equal-width count or spacing.`);
    }
    if (!Array.isArray(value.definitions) || value.definitions.length < 1 || value.definitions.length > MAX_COLUMNS) {
      throw new TypeError(`${label}.definitions must contain 1 through ${MAX_COLUMNS} columns.`);
    }
    if (!Object.hasOwn(value, "separator") || typeof value.separator !== "boolean") {
      throw new TypeError(`${label}.separator must be an explicit boolean.`);
    }
    const definitions = value.definitions.map((definition, index) => {
      if (!definition || typeof definition !== "object" || Array.isArray(definition)) {
        throw new TypeError(`${label}.definitions[${index}] must be an object.`);
      }
      const definitionUnsupported = Object.keys(definition).filter((key) => !COLUMN_DEFINITION_KEYS.has(key));
      if (definitionUnsupported.length) {
        throw new TypeError(`${label}.definitions[${index}] has unsupported properties: ${definitionUnsupported.join(", ")}.`);
      }
      if (!Object.hasOwn(definition, "width") || !Object.hasOwn(definition, "spacing")) {
        throw new TypeError(`${label}.definitions[${index}] requires width and spacing.`);
      }
      return {
        width: normalizedPositiveInteger(definition.width, `${label}.definitions[${index}].width`, MAX_TWIPS),
        spacing: normalizedNonnegativeInteger(definition.spacing, `${label}.definitions[${index}].spacing`, MAX_TWIPS),
      };
    });
    return { definitions, separator: value.separator };
  }
  for (const key of ["count", "spacing", "separator"]) {
    if (!Object.hasOwn(value, key)) throw new TypeError(`${label}.${key} is required.`);
  }
  if (typeof value.separator !== "boolean") throw new TypeError(`${label}.separator must be a boolean.`);
  return {
    count: normalizedPositiveInteger(value.count, `${label}.count`, MAX_COLUMNS),
    spacing: normalizedNonnegativeInteger(value.spacing, `${label}.spacing`, MAX_TWIPS),
    separator: value.separator,
  };
}

function canonicalUnsignedInteger(value, label, maximum, { positive = false } = {}) {
  if (!/^(?:0|[1-9]\d*)$/.test(String(value))) throw new Error(`${label} must be a canonical unsigned integer.`);
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed) || parsed > maximum || (positive && parsed < 1)) {
    throw new Error(`${label} must be an integer from ${positive ? 1 : 0} through ${maximum}.`);
  }
  return parsed;
}

function canonicalBoolean(value, label) {
  if (value === "true" || value === "1") return true;
  if (value === "false" || value === "0") return false;
  throw new Error(`${label} must be true, false, 1, or 0.`);
}

function onlyAttributes(attributes, allowed, label) {
  const unknown = Object.keys(attributes).filter((key) => !allowed.has(key));
  if (unknown.length) throw new Error(`${label} has unsupported attributes: ${unknown.join(", ")}.`);
}

function rootColumnsAttributes(openingTag, label) {
  if (!/^<w:cols\b[^>]*>$/.test(openingTag)) throw new Error(`${label} is not a canonical w:cols opening tag.`);
  return wordAttributes(`${openingTag.slice(0, -1)}/>`, label);
}

function canonicalColumnsMarkup(sectionXml, label) {
  const matches = [...String(sectionXml).matchAll(/<w:cols\b[^>]*?\/>|<w:cols\b[^>]*>[\s\S]*?<\/w:cols>/g)];
  if (matches.length !== 1) throw new Error(`${label} must contain exactly one canonical w:cols element; found ${matches.length}.`);
  const markup = matches[0][0];
  const selfClosing = /^<w:cols\b[^>]*\/>$/.test(markup);
  const paired = /^((?:<w:cols\b[^>]*>))([\s\S]*)<\/w:cols>$/.exec(markup);
  const openingTag = selfClosing ? markup : paired?.[1];
  if (!openingTag) throw new Error(`${label} has unsupported w:cols markup.`);
  const attributes = selfClosing ? wordAttributes(markup, label) : rootColumnsAttributes(openingTag, label);
  const equalWidth = Object.hasOwn(attributes, "equalWidth")
    ? canonicalBoolean(attributes.equalWidth, `${label} w:equalWidth`)
    : true;
  if (equalWidth) {
    if (!selfClosing) throw new Error(`${label} equal-width w:cols cannot contain child elements.`);
    onlyAttributes(attributes, EQUAL_WIDTH_ATTRIBUTES, `${label} equal-width w:cols`);
    if (!Object.hasOwn(attributes, "space")) throw new Error(`${label} equal-width w:cols is missing w:space.`);
    return {
      markup,
      columns: {
        count: Object.hasOwn(attributes, "num")
          ? canonicalUnsignedInteger(attributes.num, `${label} w:num`, MAX_COLUMNS, { positive: true })
          : 1,
        spacing: canonicalUnsignedInteger(attributes.space, `${label} w:space`, MAX_TWIPS),
        separator: Object.hasOwn(attributes, "sep") ? canonicalBoolean(attributes.sep, `${label} w:sep`) : false,
      },
    };
  }
  if (selfClosing || !paired) throw new Error(`${label} custom-width w:cols requires canonical w:col children.`);
  onlyAttributes(attributes, CUSTOM_WIDTH_ATTRIBUTES, `${label} custom-width w:cols`);
  const inner = paired[2];
  const columnMatches = [...inner.matchAll(/<w:col\b[^>]*\/>/g)];
  if (columnMatches.length < 1 || columnMatches.length > MAX_COLUMNS) {
    throw new Error(`${label} custom-width w:cols requires 1 through ${MAX_COLUMNS} canonical w:col leaves.`);
  }
  if (inner.replace(/<w:col\b[^>]*\/>/g, "").trim()) {
    throw new Error(`${label} custom-width w:cols has unsupported child content.`);
  }
  const definitions = columnMatches.map((match, index) => {
    const definition = wordAttributes(match[0], `${label} w:col ${index}`);
    onlyAttributes(definition, CUSTOM_COLUMN_ATTRIBUTES, `${label} w:col ${index}`);
    if (!Object.hasOwn(definition, "w")) throw new Error(`${label} w:col ${index} is missing w:w.`);
    return {
      width: canonicalUnsignedInteger(definition.w, `${label} w:col ${index} w:w`, MAX_TWIPS, { positive: true }),
      spacing: Object.hasOwn(definition, "space")
        ? canonicalUnsignedInteger(definition.space, `${label} w:col ${index} w:space`, MAX_TWIPS)
        : 0,
    };
  });
  return {
    markup,
    columns: {
      definitions,
      separator: Object.hasOwn(attributes, "sep") ? canonicalBoolean(attributes.sep, `${label} w:sep`) : false,
    },
  };
}

function rawSectionColumns(xml, sectionOrdinal, label) {
  const sections = sectionProperties(xml);
  if (!sections[sectionOrdinal]) throw new Error(`${label} is outside the native w:sectPr sequence.`);
  return { sections, section: sections[sectionOrdinal], ...canonicalColumnsMarkup(sections[sectionOrdinal].xml, label) };
}

function normalizeTargetColumnsXml(xml, sectionOrdinal, label) {
  const raw = rawSectionColumns(xml, sectionOrdinal, label);
  const normalizedSection = raw.section.xml.replace(raw.markup, "<w:cols officeKitColumnsMasked=\"true\"/>");
  return {
    columns: raw.columns,
    normalized: canonicalizeXmlForResidual(`${xml.slice(0, raw.section.offset)}${normalizedSection}${xml.slice(raw.section.offset + raw.section.xml.length)}`, label),
    sectionCount: raw.sections.length,
  };
}

function selectSection(document, { sectionBlockIndex, expectedColumns }) {
  const selected = selectCanonicalSection(document, sectionBlockIndex);
  const actual = normalizeColumns(selected.block.columns, "selected section columns");
  if (!equalJson(actual, expectedColumns)) {
    throw new Error(`Selected section columns do not match the expected source value: expected ${JSON.stringify(expectedColumns)}, observed ${JSON.stringify(actual)}.`);
  }
  return selected;
}

function assertValidReplacement(document) {
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Replacement section columns fail document verification: ${verification.ndjson}`);
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Changes one imported canonical w:cols element through the public DocumentFile
 * path. The complete equal-width or custom-width profile is atomic: this does
 * not convert structure, add/remove the element, or calculate reflow.
 */
export async function editImportedSectionColumns({
  inputPath,
  outputPath,
  auditPath,
  sectionBlockIndex,
  expectedColumns,
  replacementColumns,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const expected = normalizeColumns(expectedColumns, "expectedColumns");
  const replacement = normalizeColumns(replacementColumns, "replacementColumns");
  if (equalJson(expected, replacement)) throw new Error("replacementColumns must differ from expectedColumns.");
  if (Object.hasOwn(expected, "definitions") !== Object.hasOwn(replacement, "definitions")) {
    throw new Error("replacementColumns must retain the source equal-width or custom-width profile.");
  }
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectSection(document, { sectionBlockIndex, expectedColumns: expected });
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizeTargetColumnsXml(sourceXml, selected.sectionOrdinal, "source target section");
  if (!equalJson(sourceResidual.columns, expected)) {
    throw new Error("The raw source w:cols does not match the inspected section columns.");
  }
  const beforeSections = sectionProjection(document);
  selected.block.columns = structuredClone(replacement);
  assertValidReplacement(document);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound section columns edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound section columns edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }

    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizeTargetColumnsXml(outputXml, selected.sectionOrdinal, "output target section");
    if (!equalJson(outputResidual.columns, replacement)) {
      throw new Error("Exported target w:cols does not match the requested section-column replacement.");
    }
    if (sourceResidual.sectionCount !== outputResidual.sectionCount || outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Section columns edit changed word/document.xml outside the one requested canonical w:cols element.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectSection(reimported, { sectionBlockIndex: selected.blockIndex, expectedColumns: replacement });
    const afterSections = sectionProjection(reimported);
    const expectedSections = structuredClone(beforeSections);
    const expectedSection = expectedSections.find((section) => section.id === selected.snapshot.id);
    if (!expectedSection) throw new Error("Selected section disappeared from the imported section projection.");
    expectedSection.columns = structuredClone(replacement);
    if (!equalJson(afterSections, expectedSections)) {
      throw new Error("DOCX export changed imported section identity or settings outside the requested columns element.");
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
        type: "source-bound-section-columns-edit",
        target: {
          id: selected.snapshot.id,
          blockIndex: selected.blockIndex,
          sectionOrdinal: selected.sectionOrdinal,
        },
        sourceColumns: expected,
        replacementColumns: replacement,
      },
      validation: {
        changedParts: changed,
        columnsXmlResidual: {
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
          columns: replacement,
        },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes only one selected section's canonical w:cols element. It does not calculate or visually infer Word column flow; inspect a native Word or LibreOffice render before delivery."],
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

function parseJsonColumns(value, label) {
  try {
    return JSON.parse(requiredText(value, label));
  } catch (error) {
    throw new TypeError(`${label} must be JSON for a columns object: ${error.message}`);
  }
}

export function parseSectionColumnsEditCli(argv) {
  const [inputPath, outputPath, auditPath, sectionBlockIndex, expectedColumns, replacementColumns] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    sectionBlockIndex: boundedIndex(sectionBlockIndex, "sectionBlockIndex"),
    expectedColumns: parseJsonColumns(expectedColumns, "expectedColumns"),
    replacementColumns: parseJsonColumns(replacementColumns, "replacementColumns"),
  };
}

export function sectionColumnsCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedSectionColumns(parseSectionColumnsEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(sectionColumnsCliOutput(result)));
}
