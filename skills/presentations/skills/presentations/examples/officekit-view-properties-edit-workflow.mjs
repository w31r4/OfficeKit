import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "office-kit";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const VIEW_PROPERTIES_PART = "ppt/viewProps.xml";
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

function requiredPatch(value) {
  if (!value || typeof value !== "object" || Array.isArray(value) || !Object.keys(value).length) {
    throw new TypeError("patch must be a non-empty source-properties object.");
  }
  return value;
}

function xmlAttributes(tag) {
  const attributes = Object.create(null);
  for (const match of String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g)) attributes[match[1]] = match[3];
  return attributes;
}

function guideVisibility(xml) {
  const tag = String(xml).match(/<(?:[A-Za-z_][\w.-]*:)?cSldViewPr\b[^>]*>/i)?.[0];
  if (!tag) throw new Error("View-properties XML is missing p:cSldViewPr.");
  const attributes = xmlAttributes(tag);
  return Object.hasOwn(attributes, "showGuides") ? attributes.showGuides : undefined;
}

function viewSnapshot(view) {
  return {
    ...(view.gridSpacingCxEmu === undefined ? {} : { gridSpacingCxEmu: view.gridSpacingCxEmu }),
    ...(view.gridSpacingCyEmu === undefined ? {} : { gridSpacingCyEmu: view.gridSpacingCyEmu }),
    ...(view.slideViewSnapToGrid === undefined ? {} : { slideViewSnapToGrid: view.slideViewSnapToGrid }),
    ...(view.slideViewSnapToObjects === undefined ? {} : { slideViewSnapToObjects: view.slideViewSnapToObjects }),
    slideGuides: view.slideGuides.map((guide) => ({ ...guide })),
  };
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
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
  if (JSON.stringify(sourcePaths) !== JSON.stringify(outputPaths)) {
    throw new Error("View-properties edit changed PPTX package topology.");
  }
  const sourceView = await sourceZip.file(VIEW_PROPERTIES_PART)?.async("text");
  const outputView = await outputZip.file(VIEW_PROPERTIES_PART)?.async("text");
  if (!sourceView || !outputView) throw new Error("PPTX is missing its required imported view-properties part.");
  if (Buffer.from(sourceView).equals(Buffer.from(outputView))) {
    throw new Error("View-properties edit produced no native part change; refuse to publish a no-op request.");
  }
  for (const partPath of sourcePaths) {
    if (sourceZip.files[partPath].dir || partPath === VIEW_PROPERTIES_PART) continue;
    const before = await sourceZip.file(partPath).async("uint8array");
    const after = await outputZip.file(partPath).async("uint8array");
    if (!Buffer.from(before).equals(Buffer.from(after))) {
      throw new Error("View-properties edit changed non-target package part " + partPath + ".");
    }
  }
  if (guideVisibility(sourceView) !== guideVisibility(outputView)) {
    throw new Error("View-properties edit changed p:cSldViewPr/@showGuides, which is local editor state and must stay source-owned.");
  }
  return {
    partCount: sourcePaths.length,
    targetPart: VIEW_PROPERTIES_PART,
    nonTargetPartsByteIdentical: true,
    guideVisibilityPreserved: true,
    sourceViewPropertiesSha256: sha256(Buffer.from(sourceView)),
    outputViewPropertiesSha256: sha256(Buffer.from(outputView)),
  };
}

export async function editPptxViewProperties({ inputPath, outputPath, auditPath, patch }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const requestedPatch = requiredPatch(patch);
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) {
    throw new Error("auditPath must be distinct from the source and PPTX output paths.");
  }

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, {
    type: PPTX_MIME,
    name: path.basename(sourcePath),
  }));
  const capability = presentation.view.capability;
  if (!capability.sourceBound || !capability.partPresent || !capability.editable) {
    throw new Error("Imported PPTX view-properties part does not satisfy the fixed-topology editable profile.");
  }
  const sourceSnapshot = viewSnapshot(presentation.view);
  const sourceRenderHashes = await slideRenderHashes(presentation);
  presentation.view.setSourceProperties(requestedPatch);
  const requestedSnapshot = viewSnapshot(presentation.view);
  if (sameJson(sourceSnapshot, requestedSnapshot)) {
    throw new Error("patch must change at least one imported view-properties value; no output was written.");
  }

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
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
    if (!sameJson(viewSnapshot(reimported.view), requestedSnapshot)) {
      throw new Error("PPTX second import did not retain the requested view-properties semantics.");
    }
    if (!reimported.view.capability.editable) {
      throw new Error("PPTX second import did not retain an editable fixed-topology view-properties profile.");
    }
    const outputRenderHashes = await slideRenderHashes(reimported);
    if (!sameJson(sourceRenderHashes, outputRenderHashes)) {
      throw new Error("View-properties edit changed a slide model render even though view metadata is not slide content.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    const audit = {
      schema: "office-kit.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "source-bound-view-properties-edit",
        partPath: VIEW_PROPERTIES_PART,
        requestedPatch,
        sourceSnapshot,
        outputSnapshot: requestedSnapshot,
        capability,
      },
      warnings: [],
      validation: {
        package: { ok: true, ...packageScope, onlyViewPropertiesPartChanged: true },
        reimport: { ok: true, fixedTopologyEditable: reimported.view.capability.editable, requestedSemanticsRetained: true },
        modelRender: { ok: true, sourceSha256: sourceRenderHashes, outputSha256: outputRenderHashes, byteIdentical: true },
        verify: { ok: verification.ok },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await fs.rename(temporaryPath, finalPath);
    await fs.rename(temporaryAuditPath, finalAuditPath);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([
      fs.rm(temporaryPath, { force: true }),
      fs.rm(temporaryAuditPath, { force: true }),
    ]);
    throw error;
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, patchJson] = argv;
  let patch;
  try {
    patch = JSON.parse(requiredText(patchJson, "patchJson"));
  } catch (error) {
    throw new TypeError("patchJson must be one valid JSON object: " + error.message);
  }
  return { inputPath, outputPath, auditPath, patch };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxViewProperties(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    targetPart: result.audit.operation.partPath,
  }));
}
