import crypto from "node:crypto";
import { constants as FS_CONSTANTS } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";

import { FileBlob, SpreadsheetFile } from "office-kit";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const ACCESSIBILITY_FIELDS = ["title", "description", "decorative"];
const DRAWING_KINDS = new Set(["image", "chart"]);
const PACKAGE_LIMITS = {
  maxInputBytes: 64 * 1024 * 1024,
  maxParts: 5_000,
  maxPartBytes: 32 * 1024 * 1024,
  maxTotalBytes: 256 * 1024 * 1024,
  verifyCrc32: true,
  maxChars: 200_000,
};
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

function isXmlSafeText(value) {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code === 0x09 || code === 0x0a || code === 0x0d) continue;
    if (code < 0x20 || code === 0x7f || code === 0xfffe || code === 0xffff) return false;
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) return false;
  }
  return true;
}

function accessibilityState(value, label, { update = false } = {}) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError(`${label} must be an object with title, description, and/or decorative.`);
  }
  const unknown = Object.keys(value).filter((field) => !ACCESSIBILITY_FIELDS.includes(field));
  if (unknown.length) throw new TypeError(`${label} does not support ${unknown.join(", ")}.`);
  if (update && !ACCESSIBILITY_FIELDS.some((field) => Object.hasOwn(value, field))) {
    throw new TypeError(`${label} must change title, description, and/or decorative.`);
  }
  const normalized = {};
  for (const field of ACCESSIBILITY_FIELDS) {
    if (!Object.hasOwn(value, field)) continue;
    const fieldValue = value[field];
    if (fieldValue == null) {
      if (!update) throw new TypeError(`${label}.${field} cannot be null in an expected complete state; omit the field instead.`);
      normalized[field] = null;
    } else if (field === "decorative") {
      if (typeof fieldValue !== "boolean") throw new TypeError(`${label}.decorative must be a boolean.`);
      normalized.decorative = fieldValue;
    } else {
      if (typeof fieldValue !== "string" || !fieldValue.length || fieldValue.length > 1_024 || !isXmlSafeText(fieldValue)) {
        throw new TypeError(`${label}.${field} must contain 1 through 1024 XML-safe characters.`);
      }
      normalized[field] = fieldValue;
    }
  }
  if (!update && normalized.decorative === true && (normalized.title !== undefined || normalized.description !== undefined)) {
    throw new TypeError(`${label} cannot combine decorative: true with title or description.`);
  }
  return normalized;
}

function canonicalAccessibility(value) {
  const state = value || {};
  return Object.fromEntries(ACCESSIBILITY_FIELDS.filter((field) => Object.hasOwn(state, field)).map((field) => [field, state[field]]));
}

async function packageVersion() {
  const entry = require.resolve("office-kit");
  return JSON.parse(await fs.readFile(path.join(path.dirname(path.dirname(entry)), "package.json"), "utf8")).version;
}

async function assertAbsent(filePath, label) {
  try { await fs.lstat(filePath); }
  catch (error) { if (error?.code === "ENOENT") return; throw error; }
  throw new Error(`${label} already exists; refusing to overwrite it.`);
}

async function publishNoReplace(temporaryPath, finalPath, label) {
  try { await fs.link(temporaryPath, finalPath); }
  catch (error) {
    if (error?.code === "EEXIST") throw new Error(`${label} already exists; refusing to overwrite it.`);
    if (!["EPERM", "EXDEV", "ENOTSUP", "EOPNOTSUPP"].includes(error?.code)) throw error;
    try { await fs.copyFile(temporaryPath, finalPath, FS_CONSTANTS.COPYFILE_EXCL); }
    catch (copyError) {
      if (copyError?.code === "EEXIST") throw new Error(`${label} already exists; refusing to overwrite it.`);
      throw copyError;
    }
  }
  await fs.rm(temporaryPath, { force: true });
}

async function writePrivateFile(filePath, bytes) {
  const handle = await fs.open(filePath, "wx", 0o600);
  try { await handle.writeFile(bytes); await handle.sync(); }
  finally { await handle.close(); }
}

async function packageProfile(bytes, label) {
  const inspection = await SpreadsheetFile.inspectXlsx(new FileBlob(bytes, { type: XLSX_MIME }), PACKAGE_LIMITS);
  if (!inspection.ok) throw new Error(`${label} failed bounded XLSX package inspection: ${inspection.issues.map((issue) => issue.type).join(", ")}.`);
  // inspectXlsx applies declared and actual decompression budgets before this
  // second pass hashes the already-validated immutable byte buffer.
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((partPath) => !zip.files[partPath].dir).sort();
  const hashes = {};
  for (const partPath of paths) hashes[partPath] = sha256(await zip.file(partPath).async("uint8array"));
  const packageRecord = inspection.records.find((record) => record.kind === "xlsxPackage");
  return {
    paths,
    hashes,
    inspection: {
      ok: true,
      parts: packageRecord?.parts ?? paths.length,
      uncompressedBytes: packageRecord?.uncompressedBytes,
      crc32Verified: true,
    },
  };
}

function changedParts(source, output) {
  return [...new Set([...source.paths, ...output.paths])]
    .sort()
    .filter((partPath) => source.hashes[partPath] !== output.hashes[partPath]);
}

function drawingsOnSheet(sheet, kind) {
  return kind === "image" ? sheet.images.items : sheet.charts.items;
}

function selectDrawing(workbook, { sheetName, objectKind, objectName, expectedAccessibility }) {
  const sheetMatches = workbook.worksheets.items.filter((sheet) => sheet.name === sheetName);
  if (sheetMatches.length !== 1) throw new Error(`Expected exactly one worksheet named ${JSON.stringify(sheetName)}; found ${sheetMatches.length}.`);
  const sheet = sheetMatches[0];
  const matches = drawingsOnSheet(sheet, objectKind).filter((drawing) => drawing.name === objectName);
  if (matches.length !== 1) throw new Error(`Expected exactly one ${objectKind} named ${JSON.stringify(objectName)} on worksheet ${JSON.stringify(sheetName)}; found ${matches.length}.`);
  const drawing = matches[0];
  const actual = canonicalAccessibility(drawing.accessibility);
  if (!sameJson(actual, expectedAccessibility)) {
    throw new Error(`Selected ${objectKind} accessibility metadata does not match the expected complete source state: expected ${JSON.stringify(expectedAccessibility)}, observed ${JSON.stringify(actual)}.`);
  }
  if (drawing.accessibilityCapability?.sourceBound !== true || drawing.accessibilityCapability?.editable !== true) {
    throw new Error(`Selected ${objectKind} accessibility metadata is not an editable source-bound xdr:cNvPr profile.`);
  }
  if (workbook.resolve(drawing.id) !== drawing) throw new Error(`Selected ${objectKind} locator did not resolve back to the imported object.`);
  return { sheet, drawing };
}

function drawingProjection(workbook, selected) {
  return workbook.worksheets.items.map((sheet) => ({
    id: sheet.id,
    name: sheet.name,
    images: sheet.images.items.map((image) => {
      const record = structuredClone(image.toJSON());
      if (sheet.name === selected.sheetName && selected.objectKind === "image" && image.name === selected.objectName) {
        record.alt = "__OFFICEKIT_ACCESSIBILITY_MASK__";
        record.accessibility = "__OFFICEKIT_ACCESSIBILITY_MASK__";
      }
      return record;
    }),
    charts: sheet.charts.items.map((chart) => {
      const record = structuredClone(chart.toJSON());
      if (sheet.name === selected.sheetName && selected.objectKind === "chart" && chart.name === selected.objectName) {
        record.accessibility = "__OFFICEKIT_ACCESSIBILITY_MASK__";
      }
      return record;
    }),
  }));
}

async function renderHashes(workbook) {
  const output = [];
  for (const sheet of workbook.worksheets.items) {
    const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", format: "svg" });
    const svg = await preview.text();
    if (!/<svg\b/i.test(svg)) throw new Error(`Model render for worksheet ${sheet.name} did not produce SVG.`);
    const visualSvg = svg
      .replace(/<title>[\s\S]*?<\/title>/g, "")
      .replace(/<desc>[\s\S]*?<\/desc>/g, "")
      .replace(/\s+(?:role|aria-hidden)="[^"]*"/g, "");
    output.push({ sheet: sheet.name, renderer: "model-svg", semanticSha256: sha256(preview.bytes), visualSha256: sha256(Buffer.from(visualSvg)), bytes: preview.bytes.length });
  }
  return output;
}

function visualRenderProjection(records) {
  return records.map(({ sheet, renderer, visualSha256 }) => ({ sheet, renderer, visualSha256 }));
}

/**
 * Changes the accessibility title, description, or decorative classification
 * on exactly one recognized imported worksheet image or chart. The source
 * workbook, complete prior state, unique object name, and source-bound codec
 * capability are preconditions; exactly one worksheet-drawing part may differ.
 */
export async function editImportedDrawingAccessibility({
  inputPath,
  outputPath,
  auditPath,
  sheetName,
  objectKind,
  objectName,
  expectedAccessibility,
  update,
} = {}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const selected = {
    sheetName: requiredText(sheetName, "sheetName"),
    objectKind: requiredText(objectKind, "objectKind").toLowerCase(),
    objectName: requiredText(objectName, "objectName"),
  };
  if (!DRAWING_KINDS.has(selected.objectKind)) throw new TypeError("objectKind must be image or chart.");
  selected.expectedAccessibility = accessibilityState(expectedAccessibility, "expectedAccessibility");
  const requestedUpdate = accessibilityState(update, "update", { update: true });
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const sourceStat = await fs.lstat(sourcePath);
  if (!sourceStat.isFile() || sourceStat.isSymbolicLink()) throw new Error("inputPath must be a regular, non-symlink XLSX file.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceSha256 = sha256(source);
  const sourcePackage = await packageProfile(source, "Source XLSX");
  const workbook = await SpreadsheetFile.importXlsx(new FileBlob(source, { type: XLSX_MIME, name: path.basename(sourcePath) }));
  const sourceSelection = selectDrawing(workbook, selected);
  const sourceProjection = drawingProjection(workbook, selected);
  const sourceRenders = await renderHashes(workbook);
  const selectedId = sourceSelection.drawing.id;

  sourceSelection.drawing.setAccessibilityMetadata(requestedUpdate);
  const replacementAccessibility = canonicalAccessibility(sourceSelection.drawing.accessibility);
  if (sameJson(replacementAccessibility, selected.expectedAccessibility)) {
    throw new Error("update does not change the selected drawing accessibility state.");
  }
  if (replacementAccessibility.decorative !== true && replacementAccessibility.title === undefined && replacementAccessibility.description === undefined) {
    throw new Error("update must leave the selected drawing explicitly decorative or with an accessibility title/description.");
  }
  if (!sameJson(drawingProjection(workbook, selected), sourceProjection)) {
    throw new Error("The in-memory transaction changed drawing state outside the selected accessibility metadata.");
  }

  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  const nonce = `${process.pid}.${crypto.randomBytes(8).toString("hex")}`;
  const temporaryPath = path.join(path.dirname(finalPath), `.${path.basename(finalPath)}.${nonce}.tmp`);
  const temporaryAuditPath = path.join(path.dirname(finalAuditPath), `.${path.basename(finalAuditPath)}.${nonce}.tmp`);
  try {
    const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
    const output = Buffer.from(await exported.arrayBuffer());
    await writePrivateFile(temporaryPath, output);
    const outputPackage = await packageProfile(output, "Output XLSX");
    const packageChanges = changedParts(sourcePackage, outputPackage);
    if (packageChanges.length !== 1 || !/^xl\/drawings\/drawing[^/]*\.xml$/i.test(packageChanges[0])) {
      throw new Error(`Drawing accessibility edit changed unexpected package parts: ${packageChanges.join(", ") || "none"}.`);
    }

    const reimported = await SpreadsheetFile.importXlsx(new FileBlob(output, { type: XLSX_MIME, name: path.basename(finalPath) }));
    const outputSelection = selectDrawing(reimported, { ...selected, expectedAccessibility: replacementAccessibility });
    if (outputSelection.drawing.id !== selectedId) throw new Error("Second import changed the selected drawing identity.");
    if (!sameJson(drawingProjection(reimported, selected), sourceProjection)) {
      throw new Error("Second import changed modeled drawing state outside the selected accessibility metadata.");
    }
    const outputRenders = await renderHashes(reimported);
    if (!sameJson(visualRenderProjection(outputRenders), visualRenderProjection(sourceRenders))) {
      throw new Error(`Accessibility-only edit changed the workbook model render: source=${JSON.stringify(sourceRenders)} output=${JSON.stringify(outputRenders)}.`);
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Workbook verification failed after drawing accessibility edit: ${verification.ndjson}`);
    if (sha256(await fs.readFile(sourcePath)) !== sourceSha256) throw new Error("Source XLSX changed during the transaction; refusing to publish output.");

    const audit = {
      schema: "office-kit.xlsx-drawing-accessibility-edit-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceSha256, bytes: source.length, immutable: true },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { requested: "office-kit", actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", sourceMutation: false, outputCollision: "fail" },
      operation: {
        type: "drawing-accessibility-edit",
        locator: { sheet: selected.sheetName, objectKind: selected.objectKind, name: selected.objectName, id: selectedId },
        previous: selected.expectedAccessibility,
        update: requestedUpdate,
        result: replacementAccessibility,
      },
      validation: {
        sourceUnchanged: true,
        capability: { sourceBound: true, editable: true },
        package: { changedPaths: packageChanges, exactlyOneWorksheetDrawingPartChanged: true, allOtherPartsByteIdentical: true, source: sourcePackage.inspection, output: outputPackage.inspection },
        model: { nonTargetDrawingStatePreserved: true, selectedIdentityPreserved: true },
        reimport: { ok: true, accessibility: replacementAccessibility },
        verify: { ok: verification.ok, issueCount: verification.issues.length },
        modelRender: { visuallyUnchanged: true, source: sourceRenders, output: outputRenders },
      },
      boundaries: {
        mutation: "one-recognized-source-bound-xdr-cNvPr-accessibility-state",
        workbookAccessibilityConformance: false,
        readingOrderAndWorksheetIntent: "manual-native-host-or-author-review",
        note: "This transaction proves one modeled image/chart accessibility edit and part-level package locality. It is not Excel Accessibility Checker, WCAG, PDF, or whole-workbook conformance evidence.",
      },
    };
    await writePrivateFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`);
    await publishNoReplace(temporaryPath, finalPath, "outputPath");
    try { await publishNoReplace(temporaryAuditPath, finalAuditPath, "auditPath"); }
    catch (error) { await fs.rm(finalPath, { force: true }); throw error; }
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } finally {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
  }
}

function parseJson(value, label) {
  try { return JSON.parse(requiredText(value, label)); }
  catch (error) { throw new TypeError(`${label} must be valid JSON: ${error.message}`); }
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const [inputPath, outputPath, auditPath, sheetName, objectKind, objectName, expectedJson, updateJson] = process.argv.slice(2);
  const result = await editImportedDrawingAccessibility({
    inputPath,
    outputPath,
    auditPath,
    sheetName,
    objectKind,
    objectName,
    expectedAccessibility: parseJson(expectedJson, "expectedAccessibilityJson"),
    update: parseJson(updateJson, "updateJson"),
  });
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256, changedPaths: result.audit.validation.package.changedPaths }));
}
