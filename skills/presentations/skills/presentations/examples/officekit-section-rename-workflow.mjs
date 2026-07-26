import crypto from "node:crypto";
import { constants as FS_CONSTANTS } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "office-kit";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_SECTION_NAME = 255;
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("office-kit");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty string.");
  return value.trim();
}

function sectionName(value, label) {
  const name = requiredText(value, label);
  if (name.length > MAX_SECTION_NAME) {
    throw new RangeError(label + " must contain 1 through " + MAX_SECTION_NAME + " characters.");
  }
  if (/[\u0000-\u001f\u007f]/.test(name)) throw new TypeError(label + " must not contain control characters.");
  return name;
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
  if (!sameJson(sourcePaths, outputPaths)) throw new Error("Section rename changed PPTX package topology.");
  const sourcePresentation = await sourceZip.file("ppt/presentation.xml")?.async("uint8array");
  const outputPresentation = await outputZip.file("ppt/presentation.xml")?.async("uint8array");
  if (!sourcePresentation || !outputPresentation) throw new Error("PPTX is missing ppt/presentation.xml.");
  if (Buffer.from(sourcePresentation).equals(Buffer.from(outputPresentation))) {
    throw new Error("Section rename produced no ppt/presentation.xml change; no output was published.");
  }
  for (const partPath of sourcePaths) {
    if (sourceZip.files[partPath].dir || partPath === "ppt/presentation.xml") continue;
    const before = await sourceZip.file(partPath).async("uint8array");
    const after = await outputZip.file(partPath).async("uint8array");
    if (!Buffer.from(before).equals(Buffer.from(after))) {
      throw new Error("Section rename changed non-target package part " + partPath + ".");
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

// This is deliberately a rename-only transaction. The public section model can
// move a canonical boundary, but an Agent-facing rename should not also carry a
// hidden partition rewrite: membership has a separate, larger precondition.
export async function renamePptxSection({ inputPath, outputPath, auditPath, expectedName, replacementName }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const sourceName = sectionName(expectedName, "expectedName");
  const outputName = sectionName(replacementName, "replacementName");
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) {
    throw new Error("auditPath must be distinct from source and PPTX output paths.");
  }
  if (sourceName === outputName) throw new Error("replacementName must differ from expectedName.");
  await assertAbsent(finalPath, "outputPath");
  await assertAbsent(finalAuditPath, "auditPath");

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, {
    type: PPTX_MIME,
    name: path.basename(sourcePath),
  }));
  const sourceSections = sectionSnapshot(presentation);
  if (!sourceSections.length) {
    throw new Error("PPTX has no semantic canonical PowerPoint sections; opaque or section-free sources cannot use this rename transaction.");
  }
  const targets = presentation.sections.items.filter((section) => section.name === sourceName);
  if (targets.length !== 1) {
    throw new Error("Expected exactly one imported PowerPoint section named " + JSON.stringify(sourceName) + "; found " + targets.length + ".");
  }
  const target = targets[0];
  const targetIndex = presentation.sections.items.indexOf(target);
  if (!target.id || !target.nativeId || !Array.isArray(target.slideIds) || !target.slideIds.length) {
    throw new Error("Selected PowerPoint section is missing its fixed source identity or membership.");
  }
  const conflicting = presentation.sections.items.find((section) => section !== target && section.name.toLowerCase() === outputName.toLowerCase());
  if (conflicting) {
    throw new Error("replacementName conflicts case-insensitively with existing PowerPoint section " + JSON.stringify(conflicting.name) + ".");
  }
  const sourceNonSection = nonSectionSnapshot(presentation);
  const sourceRenderHashes = await slideRenderHashes(presentation);
  const expectedSections = structuredClone(sourceSections);
  expectedSections[targetIndex].name = outputName;
  target.name = outputName;

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  let publishedOutput = false;
  let publishedAudit = false;
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertPackageScope(source, output);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, {
      type: PPTX_MIME,
      name: path.basename(finalPath),
    }));
    const outputSections = sectionSnapshot(reimported);
    if (!sameJson(outputSections, expectedSections)) {
      throw new Error("PPTX second import did not retain exactly the requested fixed-topology section rename.");
    }
    if (!sameJson(nonSectionSnapshot(reimported), sourceNonSection)) {
      throw new Error("Section rename changed a non-section presentation semantic.");
    }
    const outputRenderHashes = await slideRenderHashes(reimported);
    if (!sameJson(sourceRenderHashes, outputRenderHashes)) {
      throw new Error("Section rename changed a static slide model render.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    if (!Buffer.from(await fs.readFile(sourcePath)).equals(source)) {
      throw new Error("Source PPTX changed during section rename transaction; refusing to publish output.");
    }
    const audit = {
      schema: "office-kit.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "source-bound-section-name-edit",
        sectionId: target.id,
        nativeId: target.nativeId,
        ordinal: targetIndex + 1,
        expectedName: sourceName,
        replacementName: outputName,
        orderedSlideIds: [...sourceSections[targetIndex].slideIds],
      },
      warnings: ["Static render verification proves visible-slide stability, not PowerPoint navigation-pane behavior."],
      validation: {
        package: { ok: true, ...packageScope, onlyPresentationPartChanged: true },
        reimport: { ok: true, sectionCount: outputSections.length, exactFixedTopologyRetained: true },
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
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, expectedName, replacementName] = argv;
  return { inputPath, outputPath, auditPath, expectedName, replacementName };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await renamePptxSection(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    targetPart: result.audit.validation.package.targetPart,
  }));
}
