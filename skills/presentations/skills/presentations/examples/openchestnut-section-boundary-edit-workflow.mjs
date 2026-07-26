import crypto from "node:crypto";
import { constants as FS_CONSTANTS } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "open-office-artifact-tool";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_SECTIONS = 4096;
const MAX_SECTION_NAME = 255;
const MAX_SECTION_SLIDES = 16384;
const MAX_PARTITION_JSON_BYTES = 32 * 1024 * 1024;
const SECTION_GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty string.");
  return value.trim();
}

function requiredExactText(value, label) {
  if (typeof value !== "string" || !value || value !== value.trim()) {
    throw new TypeError(label + " must be a non-empty string without leading or trailing whitespace.");
  }
  return value;
}

function sectionName(value, label) {
  const name = requiredExactText(value, label);
  if (name.length > MAX_SECTION_NAME) {
    throw new RangeError(label + " must contain 1 through " + MAX_SECTION_NAME + " characters.");
  }
  if (/[\u0000-\u001f\u007f]/.test(name)) throw new TypeError(label + " must not contain control characters.");
  return name;
}

function sectionGuid(value, label) {
  const nativeId = requiredExactText(value, label);
  if (!SECTION_GUID.test(nativeId)) throw new TypeError(label + " must be an uppercase brace-delimited GUID.");
  return nativeId;
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function sectionSnapshot(presentation) {
  return presentation.sections.items.map((section) => section.toJSON());
}

function nonSectionSnapshot(presentation) {
  const snapshot = presentation.toProto();
  delete snapshot.sections;
  return snapshot;
}

function normalizedPartition(value, label) {
  if (!Array.isArray(value) || value.length < 1 || value.length > MAX_SECTIONS) {
    throw new RangeError(label + " must be an array containing 1 through " + MAX_SECTIONS + " sections.");
  }
  const ids = new Set();
  const names = new Set();
  const nativeIds = new Set();
  return value.map((candidate, index) => {
    const itemLabel = label + "[" + index + "]";
    if (!candidate || typeof candidate !== "object" || Array.isArray(candidate)) {
      throw new TypeError(itemLabel + " must be a section object.");
    }
    const keys = Object.keys(candidate).sort();
    if (!sameJson(keys, ["id", "name", "nativeId", "slideIds"])) {
      throw new TypeError(itemLabel + " must contain exactly id, name, nativeId, and slideIds.");
    }
    const id = requiredExactText(candidate.id, itemLabel + ".id");
    if (id.length > 1024 || ids.has(id)) throw new Error(itemLabel + ".id must be unique and at most 1024 characters.");
    ids.add(id);
    const name = sectionName(candidate.name, itemLabel + ".name");
    const normalizedName = name.toLowerCase();
    if (names.has(normalizedName)) throw new Error(itemLabel + ".name must be unique case-insensitively.");
    names.add(normalizedName);
    const nativeId = sectionGuid(candidate.nativeId, itemLabel + ".nativeId");
    if (nativeIds.has(nativeId)) throw new Error(itemLabel + ".nativeId must be unique.");
    nativeIds.add(nativeId);
    if (!Array.isArray(candidate.slideIds) || candidate.slideIds.length < 1 || candidate.slideIds.length > MAX_SECTION_SLIDES) {
      throw new RangeError(itemLabel + ".slideIds must contain 1 through " + MAX_SECTION_SLIDES + " entries.");
    }
    const slideIds = candidate.slideIds.map((slideId, slideIndex) => {
      const value = requiredExactText(slideId, itemLabel + ".slideIds[" + slideIndex + "]");
      if (value.length > 1024) throw new RangeError(itemLabel + ".slideIds[" + slideIndex + "] exceeds 1024 characters.");
      return value;
    });
    return { id, name, nativeId, slideIds };
  });
}

function assertTargetPartition(source, expected, replacement, presentation) {
  if (!sameJson(source, expected)) {
    throw new Error("expectedSections does not exactly match the current canonical source-bound PowerPoint section partition.");
  }
  if (source.length < 2) {
    throw new Error("PowerPoint section boundary edits require at least two imported sections.");
  }
  for (let index = 0; index < source.length; index += 1) {
    const current = source[index];
    const target = replacement[index];
    if (current.id !== target.id || current.name !== target.name || current.nativeId !== target.nativeId) {
      throw new Error("replacementSections[" + index + "] must preserve the source section ID, name, and native GUID; use the separate rename workflow for a label change.");
    }
  }
  const orderedDeckSlideIds = presentation.slides.items.map((slide) => slide.id);
  const replacementMembership = replacement.flatMap((section) => section.slideIds);
  if (!sameJson(replacementMembership, orderedDeckSlideIds)) {
    throw new Error("replacementSections must partition every retained deck slide exactly once and in presentation order.");
  }
  if (sameJson(source, replacement)) {
    throw new Error("replacementSections makes no section-boundary change; no output was published.");
  }
}

async function slideRenderHashes(presentation) {
  return Promise.all(presentation.slides.items.map(async (slide) => {
    const rendered = await slide.export({ format: "svg" });
    const svg = await rendered.text();
    if (!/<svg\b/i.test(svg)) throw new Error("Presentation model render did not produce SVG.");
    return sha256(Buffer.from(svg));
  }));
}

async function assertPackageScope(sourceBytes, outputBytes) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const sourcePaths = Object.keys(sourceZip.files).sort();
  const outputPaths = Object.keys(outputZip.files).sort();
  if (!sameJson(sourcePaths, outputPaths)) throw new Error("Section boundary edit changed PPTX package topology.");
  const sourcePresentation = await sourceZip.file("ppt/presentation.xml")?.async("uint8array");
  const outputPresentation = await outputZip.file("ppt/presentation.xml")?.async("uint8array");
  if (!sourcePresentation || !outputPresentation) throw new Error("PPTX is missing ppt/presentation.xml.");
  if (Buffer.from(sourcePresentation).equals(Buffer.from(outputPresentation))) {
    throw new Error("Section boundary edit produced no ppt/presentation.xml change; no output was published.");
  }
  for (const partPath of sourcePaths) {
    if (sourceZip.files[partPath].dir || partPath === "ppt/presentation.xml") continue;
    const before = await sourceZip.file(partPath).async("uint8array");
    const after = await outputZip.file(partPath).async("uint8array");
    if (!Buffer.from(before).equals(Buffer.from(after))) {
      throw new Error("Section boundary edit changed non-target package part " + partPath + ".");
    }
  }
  return {
    targetPart: "ppt/presentation.xml",
    partCount: sourcePaths.length,
    nonTargetPartsByteIdentical: true,
    sourceTargetSha256: sha256(sourcePresentation),
    outputTargetSha256: sha256(outputPresentation),
  };
}

async function assertAbsent(filePath, label) {
  try {
    await fs.lstat(filePath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(label + " already exists; refusing to overwrite it.");
}

async function publishNoReplace(temporaryPath, finalPath, label) {
  try {
    await fs.link(temporaryPath, finalPath);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(label + " already exists; refusing to overwrite it.");
    if (!["EPERM", "EXDEV", "ENOTSUP", "EOPNOTSUPP"].includes(error?.code)) throw error;
    try {
      await fs.copyFile(temporaryPath, finalPath, FS_CONSTANTS.COPYFILE_EXCL);
    } catch (copyError) {
      if (copyError?.code === "EEXIST") throw new Error(label + " already exists; refusing to overwrite it.");
      throw copyError;
    }
  }
  await fs.rm(temporaryPath, { force: true }).catch(() => {});
}

// This is deliberately a complete-partition transaction. The low-level public
// model can call setSlides on individual sections, but an Agent must state the
// whole requested partition: otherwise a partial list would hide which adjacent
// section receives a moved boundary slide.
export async function replacePptxSectionPartition({ inputPath, outputPath, auditPath, expectedSections, replacementSections }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) {
    throw new Error("auditPath must be distinct from source and PPTX output paths.");
  }
  await assertAbsent(finalPath, "outputPath");
  await assertAbsent(finalAuditPath, "auditPath");

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, {
    type: PPTX_MIME,
    name: path.basename(sourcePath),
  }));
  const importedSectionSnapshot = sectionSnapshot(presentation);
  if (!importedSectionSnapshot.length) {
    throw new Error("PPTX has no semantic canonical PowerPoint sections; opaque or section-free sources cannot use this boundary transaction.");
  }
  const sourceSections = normalizedPartition(importedSectionSnapshot, "sourceSections");
  const expected = normalizedPartition(expectedSections, "expectedSections");
  const replacement = normalizedPartition(replacementSections, "replacementSections");
  if (expected.length !== sourceSections.length || replacement.length !== sourceSections.length) {
    throw new Error("expectedSections and replacementSections must each list every fixed source section in source order.");
  }
  assertTargetPartition(sourceSections, expected, replacement, presentation);

  const sourceNonSection = nonSectionSnapshot(presentation);
  const sourceRenderHashes = await slideRenderHashes(presentation);
  const changedSections = sourceSections.map((section, index) => ({
    section,
    replacement: replacement[index],
    ordinal: index + 1,
  })).filter(({ section, replacement: target }) => !sameJson(section.slideIds, target.slideIds));
  for (const { replacement: target, ordinal } of changedSections) {
    presentation.sections.items[ordinal - 1].setSlides(target.slideIds);
  }

  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  const temporaryDirectory = await fs.mkdtemp(path.join(path.dirname(finalPath), ".open-office-artifact-tool-section-boundary-"));
  const temporaryPath = path.join(temporaryDirectory, "output.pptx");
  const temporaryAuditPath = path.join(temporaryDirectory, "audit.json");
  let publishedOutput = false;
  let publishedAudit = false;
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertPackageScope(source, output);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, {
      type: PPTX_MIME,
      name: path.basename(finalPath),
    }));
    const outputSections = normalizedPartition(sectionSnapshot(reimported), "outputSections");
    if (!sameJson(outputSections, replacement)) {
      throw new Error("PPTX second import did not retain exactly the requested fixed-identity section partition.");
    }
    if (!sameJson(nonSectionSnapshot(reimported), sourceNonSection)) {
      throw new Error("Section boundary edit changed a non-section presentation semantic.");
    }
    const outputRenderHashes = await slideRenderHashes(reimported);
    if (!sameJson(sourceRenderHashes, outputRenderHashes)) {
      throw new Error("Section boundary edit changed a static slide model render.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    if (!Buffer.from(await fs.readFile(sourcePath)).equals(source)) {
      throw new Error("Source PPTX changed during section boundary transaction; refusing to publish output.");
    }
    const audit = {
      schema: "open-office-artifact-tool.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "source-bound-section-boundary-edit",
        fixedSectionCount: sourceSections.length,
        expectedSections: expected,
        replacementSections: replacement,
        changedSections: changedSections.map(({ section, replacement: target, ordinal }) => ({
          sectionId: section.id,
          nativeId: section.nativeId,
          ordinal,
          name: section.name,
          expectedSlideIds: section.slideIds,
          replacementSlideIds: target.slideIds,
        })),
      },
      warnings: ["Static render verification proves visible-slide stability, not PowerPoint navigation-pane behavior."],
      validation: {
        package: { ok: true, ...packageScope, onlyPresentationPartChanged: true },
        reimport: {
          ok: true,
          sectionCount: outputSections.length,
          exactFixedIdentityPartitionRetained: true,
        },
        nonSectionSemantics: { ok: true, stable: true },
        modelRender: { ok: true, sourceSha256: sourceRenderHashes, outputSha256: outputRenderHashes, byteIdentical: true },
        verify: { ok: verification.ok },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await publishNoReplace(temporaryPath, finalPath, "outputPath");
    publishedOutput = true;
    await publishNoReplace(temporaryAuditPath, finalAuditPath, "auditPath");
    publishedAudit = true;
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([
      fs.rm(temporaryPath, { force: true }),
      fs.rm(temporaryAuditPath, { force: true }),
      ...(publishedOutput ? [fs.rm(finalPath, { force: true })] : []),
      ...(publishedAudit ? [fs.rm(finalAuditPath, { force: true })] : []),
    ]);
    throw error;
  } finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
  }
}

async function parsePartitionJson(value, label) {
  const supplied = requiredText(value, label);
  let json = supplied;
  if (supplied.startsWith("@")) {
    const filePath = path.resolve(requiredText(supplied.slice(1), label + " file path"));
    const stat = await fs.stat(filePath);
    if (!stat.isFile()) throw new TypeError(label + " file path must identify a regular file.");
    if (stat.size > MAX_PARTITION_JSON_BYTES) {
      throw new RangeError(label + " file exceeds the " + MAX_PARTITION_JSON_BYTES + "-byte JSON budget.");
    }
    json = await fs.readFile(filePath, "utf8");
  }
  try {
    return JSON.parse(json);
  } catch (error) {
    throw new TypeError(label + " must be valid JSON: " + error.message);
  }
}

async function parseCli(argv) {
  const [inputPath, outputPath, auditPath, expectedSections, replacementSections] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    expectedSections: await parsePartitionJson(expectedSections, "expectedSections JSON"),
    replacementSections: await parsePartitionJson(replacementSections, "replacementSections JSON"),
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await replacePptxSectionPartition(await parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    targetPart: result.audit.validation.package.targetPart,
    changedSections: result.audit.operation.changedSections.length,
  }));
}
