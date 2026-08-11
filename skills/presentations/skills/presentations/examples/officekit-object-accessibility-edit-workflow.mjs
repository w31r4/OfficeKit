import crypto from "node:crypto";
import { constants as FS_CONSTANTS } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";

import { FileBlob, PresentationFile } from "office-kit";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const ACCESSIBILITY_FIELDS = ["title", "description", "decorative"];
const LOCATOR_FIELDS = ["slide", "id", "objectKind", "name", "parentGroupId"];
const OBJECT_KINDS = new Set(["shape", "connector", "group", "image", "table", "chart"]);
const INSPECT_KINDS = "shape,connector,groupShape,image,table,chart";
const PACKAGE_LIMITS = {
  maxInputBytes: 64 * 1024 * 1024,
  maxParts: 5_000,
  maxPartBytes: 32 * 1024 * 1024,
  maxTotalBytes: 256 * 1024 * 1024,
  verifyCrc32: true,
  maxChars: 200_000,
};
const ACCESSIBILITY_MASK = "__OFFICEKIT_ACCESSIBILITY_MASK__";
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

function accessibilityLocator(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError("locator must be an accessibility audit locator object.");
  }
  const unknown = Object.keys(value).filter((field) => !LOCATOR_FIELDS.includes(field));
  if (unknown.length) throw new TypeError(`locator does not support ${unknown.join(", ")}.`);
  const slide = Number(value.slide);
  if (!Number.isSafeInteger(slide) || slide < 1) throw new RangeError("locator.slide must be a positive safe integer.");
  const objectKind = requiredText(value.objectKind, "locator.objectKind");
  if (!OBJECT_KINDS.has(objectKind)) throw new TypeError("locator.objectKind must be shape, connector, group, image, table, or chart.");
  const locator = {
    slide,
    id: requiredText(value.id, "locator.id"),
    objectKind,
  };
  if (Object.hasOwn(value, "name")) locator.name = requiredText(value.name, "locator.name");
  if (Object.hasOwn(value, "parentGroupId")) locator.parentGroupId = requiredText(value.parentGroupId, "locator.parentGroupId");
  return locator;
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
  await fs.rm(temporaryPath, { force: true }).catch(() => {});
}

async function writePrivateFile(filePath, bytes) {
  const handle = await fs.open(filePath, "wx", 0o600);
  try { await handle.writeFile(bytes); await handle.sync(); }
  finally { await handle.close(); }
}

function xmlAttributes(tag) {
  const attributes = Object.create(null);
  for (const match of String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g)) attributes[match[1]] = match[3];
  return attributes;
}

function resolveRelationshipTarget(target) {
  const resolved = new URL(target, "https://officekit.invalid/ppt/presentation.xml");
  if (resolved.origin !== "https://officekit.invalid") throw new Error("Unexpected PPTX relationship target origin.");
  const partPath = resolved.pathname.replace(/^\/+/, "");
  if (!partPath.startsWith("ppt/") || partPath.split("/").includes("..")) {
    throw new Error("Unsafe PPTX slide relationship target: " + JSON.stringify(target));
  }
  return partPath;
}

async function orderedSlidePartPaths(zip) {
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const relationshipsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");
  if (!presentationXml || !relationshipsXml) throw new Error("PPTX is missing presentation.xml or its relationship part.");
  const relationships = new Map();
  for (const match of relationshipsXml.matchAll(/<Relationship\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type?.endsWith("/slide")) continue;
    if (attributes.TargetMode?.toLowerCase() === "external" || !attributes.Target) {
      throw new Error(`Presentation slide relationship ${JSON.stringify(attributes.Id)} is not an internal SlidePart.`);
    }
    relationships.set(attributes.Id, resolveRelationshipTarget(attributes.Target));
  }
  const paths = [];
  for (const match of presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*>/gi)) {
    const relationshipId = xmlAttributes(match[0])["r:id"];
    const target = relationships.get(relationshipId);
    if (!target) throw new Error(`Presentation slide list references unresolved relationship ${JSON.stringify(relationshipId)}.`);
    if (!zip.file(target)) throw new Error(`Presentation slide relationship points at missing part ${target}.`);
    paths.push(target);
  }
  if (!paths.length || new Set(paths).size !== paths.length) {
    throw new Error("Presentation slide list must contain distinct, resolvable SlideParts.");
  }
  return paths;
}

async function packageProfile(bytes, label) {
  const inspection = await PresentationFile.inspectPptx(new FileBlob(bytes, { type: PPTX_MIME }), PACKAGE_LIMITS);
  if (!inspection.ok) throw new Error(`${label} failed bounded PPTX package inspection: ${inspection.issues.map((issue) => issue.type).join(", ")}.`);
  // inspectPptx applies declared and actual decompression budgets before this
  // second pass hashes the already-validated immutable byte buffer.
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((partPath) => !zip.files[partPath].dir).sort();
  const hashes = {};
  for (const partPath of paths) hashes[partPath] = sha256(await zip.file(partPath).async("uint8array"));
  const packageRecord = inspection.records.find((record) => record.kind === "pptxPackage");
  return {
    paths,
    hashes,
    slideParts: await orderedSlidePartPaths(zip),
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

function recordLocator(record) {
  const locator = {
    slide: record.slide,
    id: record.id,
    objectKind: record.kind === "groupShape" ? "group" : record.kind,
  };
  if (record.name !== undefined) locator.name = record.name;
  if (record.parentGroupId !== undefined) locator.parentGroupId = record.parentGroupId;
  return locator;
}

function selectObject(presentation, { locator, expectedAccessibility }) {
  const inspection = presentation.inspect({ kind: INSPECT_KINDS, maxChars: 2_000_000 });
  if (inspection.truncated) throw new Error("Presentation object inspection exceeded the bounded locator budget.");
  const records = inspection.ndjson.split("\n").filter(Boolean).map((line) => JSON.parse(line));
  const matches = records.filter((record) => record.id === locator.id);
  if (matches.length !== 1) throw new Error(`Expected locator.id to resolve exactly one modeled presentation object; found ${matches.length}.`);
  const record = matches[0];
  const actualLocator = recordLocator(record);
  if (!sameJson(actualLocator, locator)) {
    throw new Error(`Selected object does not match the complete audit locator: expected ${JSON.stringify(locator)}, observed ${JSON.stringify(actualLocator)}.`);
  }
  const target = presentation.resolve(locator.id);
  if (!target || typeof target.setAccessibilityMetadata !== "function") {
    throw new Error("Selected locator does not resolve to an accessibility-editable presentation object.");
  }
  const actualAccessibility = canonicalAccessibility(target.accessibility);
  if (!sameJson(actualAccessibility, expectedAccessibility)) {
    throw new Error(`Selected object accessibility metadata does not match the expected complete source state: expected ${JSON.stringify(expectedAccessibility)}, observed ${JSON.stringify(actualAccessibility)}.`);
  }
  if (target.accessibilityCapability?.sourceBound !== true || target.accessibilityCapability?.editable !== true) {
    throw new Error("Selected object accessibility metadata is not an editable source-bound p:cNvPr profile.");
  }
  return { target, record };
}

function presentationProjection(presentation, targetId, objectKind) {
  const projection = structuredClone(presentation.toProto());
  let matches = 0;
  const visit = (value) => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (!value || typeof value !== "object") return;
    if (value.id === targetId) {
      matches += 1;
      value.accessibility = ACCESSIBILITY_MASK;
      if (objectKind === "image") value.alt = ACCESSIBILITY_MASK;
    }
    for (const item of Object.values(value)) visit(item);
  };
  visit(projection);
  if (matches !== 1) throw new Error(`Presentation model projection found ${matches} objects for target ID ${JSON.stringify(targetId)}.`);
  return projection;
}

async function renderHashes(presentation) {
  const output = [];
  for (const slide of presentation.slides.items) {
    const preview = await slide.export({ format: "svg" });
    const svg = await preview.text();
    if (!/<svg\b/i.test(svg)) throw new Error(`Model render for slide ${slide.index + 1} did not produce SVG.`);
    const visualSvg = svg
      .replace(/<title>[\s\S]*?<\/title>/g, "")
      .replace(/<desc>[\s\S]*?<\/desc>/g, "")
      .replace(/\s+(?:role|aria-hidden)="[^"]*"/g, "");
    output.push({ slide: slide.index + 1, renderer: "model-svg", semanticSha256: sha256(preview.bytes), visualSha256: sha256(Buffer.from(visualSvg)), bytes: preview.bytes.length });
  }
  return output;
}

function visualRenderProjection(records) {
  return records.map(({ slide, renderer, visualSha256 }) => ({ slide, renderer, visualSha256 }));
}

/**
 * Changes the accessibility title, description, or decorative classification
 * on exactly one recognized imported PowerPoint object. The complete locator
 * and prior metadata are source preconditions; exactly its SlidePart may differ.
 */
export async function editImportedObjectAccessibility({
  inputPath,
  outputPath,
  auditPath,
  locator,
  expectedAccessibility,
  update,
} = {}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const selected = accessibilityLocator(locator);
  const expected = accessibilityState(expectedAccessibility, "expectedAccessibility");
  const requestedUpdate = accessibilityState(update, "update", { update: true });
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const sourceStat = await fs.lstat(sourcePath);
  if (!sourceStat.isFile() || sourceStat.isSymbolicLink()) throw new Error("inputPath must be a regular, non-symlink PPTX file.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceSha256 = sha256(source);
  const sourcePackage = await packageProfile(source, "Source PPTX");
  if (selected.slide > sourcePackage.slideParts.length) throw new Error("locator.slide is outside the source PPTX slide list.");
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const sourceSelection = selectObject(presentation, { locator: selected, expectedAccessibility: expected });
  const sourceProjection = presentationProjection(presentation, selected.id, selected.objectKind);
  const sourceRenders = await renderHashes(presentation);

  sourceSelection.target.setAccessibilityMetadata(requestedUpdate);
  const replacementAccessibility = canonicalAccessibility(sourceSelection.target.accessibility);
  if (sameJson(replacementAccessibility, expected)) throw new Error("update does not change the selected object accessibility state.");
  if (replacementAccessibility.decorative !== true && replacementAccessibility.title === undefined && replacementAccessibility.description === undefined) {
    throw new Error("update must leave the selected object explicitly decorative or with an accessibility title/description.");
  }
  if (!sameJson(presentationProjection(presentation, selected.id, selected.objectKind), sourceProjection)) {
    throw new Error("The in-memory transaction changed presentation state outside the selected accessibility metadata.");
  }

  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  const nonce = `${process.pid}.${crypto.randomBytes(8).toString("hex")}`;
  const temporaryPath = path.join(path.dirname(finalPath), `.${path.basename(finalPath)}.${nonce}.tmp`);
  const temporaryAuditPath = path.join(path.dirname(finalAuditPath), `.${path.basename(finalAuditPath)}.${nonce}.tmp`);
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    const output = Buffer.from(await exported.arrayBuffer());
    await writePrivateFile(temporaryPath, output);
    const outputPackage = await packageProfile(output, "Output PPTX");
    if (!sameJson(sourcePackage.paths, outputPackage.paths) || !sameJson(sourcePackage.slideParts, outputPackage.slideParts)) {
      throw new Error("Object accessibility edit changed PPTX package topology or slide-part routing.");
    }
    const packageChanges = changedParts(sourcePackage, outputPackage);
    const targetPart = sourcePackage.slideParts[selected.slide - 1];
    if (packageChanges.length !== 1 || packageChanges[0] !== targetPart) {
      throw new Error(`Object accessibility edit changed unexpected package parts: ${packageChanges.join(", ") || "none"}; expected only ${targetPart}.`);
    }

    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const outputSelection = selectObject(reimported, { locator: selected, expectedAccessibility: replacementAccessibility });
    if (outputSelection.target.id !== sourceSelection.target.id) throw new Error("Second import changed the selected object identity.");
    if (!sameJson(presentationProjection(reimported, selected.id, selected.objectKind), sourceProjection)) {
      throw new Error("Second import changed modeled presentation state outside the selected accessibility metadata.");
    }
    const outputRenders = await renderHashes(reimported);
    if (!sameJson(visualRenderProjection(outputRenders), visualRenderProjection(sourceRenders))) {
      throw new Error(`Accessibility-only edit changed the presentation model render: source=${JSON.stringify(sourceRenders)} output=${JSON.stringify(outputRenders)}.`);
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Presentation verification failed after object accessibility edit: ${verification.ndjson}`);
    const accessibilityAudit = reimported.auditAccessibility({ maxChars: 200_000 });
    if (accessibilityAudit.issues.some((issue) => issue.id === selected.id)) {
      throw new Error("Presentation accessibility audit still reports a machine issue for the edited object.");
    }
    if (sha256(await fs.readFile(sourcePath)) !== sourceSha256) throw new Error("Source PPTX changed during the transaction; refusing to publish output.");

    const audit = {
      schema: "office-kit.pptx-object-accessibility-edit-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceSha256, bytes: source.length, immutable: true },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { requested: "office-kit", actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", sourceMutation: false, outputCollision: "fail" },
      operation: {
        type: "object-accessibility-edit",
        locator: selected,
        partPath: targetPart,
        previous: expected,
        update: requestedUpdate,
        result: replacementAccessibility,
      },
      validation: {
        sourceUnchanged: true,
        capability: { sourceBound: true, editable: true },
        package: { changedPaths: packageChanges, exactlyTargetSlidePartChanged: true, allOtherPartsByteIdentical: true, source: sourcePackage.inspection, output: outputPackage.inspection },
        model: { nonTargetPresentationStatePreserved: true, selectedIdentityPreserved: true },
        reimport: { ok: true, accessibility: replacementAccessibility },
        verify: { ok: verification.ok, issueCount: verification.issues.length },
        accessibilityAudit: { targetMachineIssueAbsent: true, conformanceClaimed: false, manualReviewRequired: accessibilityAudit.manualReviewRequired },
        modelRender: { visuallyUnchanged: true, source: sourceRenders, output: outputRenders },
      },
      boundaries: {
        mutation: "one-recognized-source-bound-p-cNvPr-accessibility-state",
        presentationAccessibilityConformance: false,
        readingOrderAndOpaqueObjects: "manual-native-host-or-author-review",
        note: "This transaction proves one modeled object accessibility edit and part-level package locality. It is not PowerPoint Accessibility Checker, WCAG, PDF/UA, or whole-deck conformance evidence.",
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
  const [inputPath, outputPath, auditPath, locatorJson, expectedJson, updateJson] = process.argv.slice(2);
  const result = await editImportedObjectAccessibility({
    inputPath,
    outputPath,
    auditPath,
    locator: parseJson(locatorJson, "locatorJson"),
    expectedAccessibility: parseJson(expectedJson, "expectedAccessibilityJson"),
    update: parseJson(updateJson, "updateJson"),
  });
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256, changedPaths: result.audit.validation.package.changedPaths }));
}
