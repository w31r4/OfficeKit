#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  importArtifactTool,
  padSlideNumber,
  parseArgs,
  requireArg,
  saveBlobToFile,
} from "../container_tools/artifact_tool_utils.mjs";
import { templateTargetIds, validateTemplatePlan } from "./validate_template_plan.mjs";
import {
  assertAbsent,
  assertRegularFile,
  exportBytes,
  fileBlob,
  isWithin,
  modelVisualSha256,
  pathsOverlap,
  publishDirectoryNoReplace,
  publishFileNoReplace,
  relativeFromWorkspace,
  runContactSheet,
  sha256,
  slidesFromPresentation,
  writeJson,
} from "./template_transaction_utils.mjs";

const MAX_FRAME_MAP_BYTES = 1_000_000;
const MAX_OUTPUT_SLIDES = 256;
const INSPECT_KINDS = "slide,textbox,shape,image,table,chart,connector,groupShape,nativeObject";
const MAX_INSPECT_CHARS = 20_000_000;

function usage() {
  return [
    "Usage:",
    "  officekit run template_following_scripts/prepare_template_starter_deck.mjs --workspace <dir> --pptx <source.pptx> --map <template-frame-map.json> --out <starter.pptx> [options]",
    "",
    "Options:",
    "  --preview-dir <dir>     Render starter slide PNGs. Defaults to <workspace>/template-starter-preview.",
    "  --layout-dir <dir>      Write starter layout JSON. Defaults to <workspace>/template-starter-layout.",
    "  --inspect <path>        template-inspect.ndjson. Defaults to <workspace>/template-inspect/template-inspect.ndjson.",
    "  --contact-sheet <path>  Optional PNG contact sheet path.",
    "  --scale <n>             Render scale. Defaults to 1.",
    "",
    "Executes the validated frame map through one source-bound clone per",
    "export/reimport boundary, removes the source slides, verifies the result,",
    "and publishes an immutable starter deck plus locator/provenance manifest.",
  ].join("\n");
}

function normalizeOutputSlides(map, sourceSlideCount) {
  if (!Array.isArray(map.outputSlides) || map.outputSlides.length === 0) {
    throw new Error("template-frame-map.json must include a non-empty outputSlides array.");
  }
  if (map.outputSlides.length > MAX_OUTPUT_SLIDES) {
    throw new Error(`template-frame-map.json exceeds the ${MAX_OUTPUT_SLIDES}-slide starter budget.`);
  }
  const outputSlides = [...map.outputSlides].sort((left, right) => Number(left.outputSlide) - Number(right.outputSlide));
  for (let index = 0; index < outputSlides.length; index += 1) {
    const entry = outputSlides[index];
    if (Number(entry.outputSlide) !== index + 1) {
      throw new Error(`outputSlides must be sequential from 1; expected outputSlide ${index + 1}.`);
    }
    if (!Number.isInteger(Number(entry.sourceSlide)) || Number(entry.sourceSlide) < 1 || Number(entry.sourceSlide) > sourceSlideCount) {
      throw new Error(`outputSlide ${entry.outputSlide} must reference sourceSlide 1-${sourceSlideCount}; got ${entry.sourceSlide}.`);
    }
  }
  return outputSlides;
}

function inspectRecordsBySlide(presentation) {
  const inspection = presentation.inspect({ kind: INSPECT_KINDS, maxChars: MAX_INSPECT_CHARS });
  if (inspection.truncated) throw new Error(`Starter locator inspection exceeded ${MAX_INSPECT_CHARS} characters.`);
  const bySlide = new Map();
  for (const line of String(inspection.ndjson || "").split(/\r?\n/)) {
    if (!line.trim()) continue;
    const record = JSON.parse(line);
    if (!Number.isInteger(record.slide) || record.kind === "slide") continue;
    if (!bySlide.has(record.slide)) bySlide.set(record.slide, []);
    bySlide.get(record.slide).push(record);
  }
  return bySlide;
}

function recordShape(record) {
  return JSON.stringify({
    kind: record.kind,
    name: record.name || "",
    text: record.text ?? record.textPreview ?? "",
    bbox: record.bbox || null,
    nativeKind: record.nativeKind || "",
  });
}

function translateStarterTargets(sourceRecords, outputRecords, entry) {
  if (sourceRecords.length !== outputRecords.length) {
    throw new Error(`outputSlide ${entry.outputSlide} clone inspection count differs from sourceSlide ${entry.sourceSlide}.`);
  }
  const sourceToOutput = new Map();
  const locators = [];
  for (let index = 0; index < sourceRecords.length; index += 1) {
    const source = sourceRecords[index];
    const output = outputRecords[index];
    if (recordShape(source) !== recordShape(output) || typeof source.id !== "string" || typeof output.id !== "string") {
      throw new Error(`outputSlide ${entry.outputSlide} clone inspection diverged at inherited element ${index + 1}.`);
    }
    sourceToOutput.set(source.id, output.id);
    locators.push({
      sourceElementId: source.id,
      starterElementId: output.id,
      kind: output.kind,
      ...(output.name ? { name: output.name } : {}),
    });
  }
  const editTargets = entry.editTargets.map((target, targetIndex) => {
    const sourceElementIds = templateTargetIds(target);
    const starterElementIds = sourceElementIds.map((sourceElementId) => {
      const translated = sourceToOutput.get(sourceElementId);
      if (!translated) {
        throw new Error(`outputSlide ${entry.outputSlide} editTargets[${targetIndex}] source locator ${sourceElementId} has no clone locator.`);
      }
      return translated;
    });
    return { targetIndex, action: target.action, sourceElementIds, starterElementIds };
  });
  return { locators, editTargets };
}

function manifestPathFor(out) {
  return out.toLowerCase().endsWith(".pptx") ? `${out.slice(0, -5)}.manifest.json` : `${out}.manifest.json`;
}

function assertDisjointOutputs(paths) {
  for (let left = 0; left < paths.length; left += 1) {
    for (let right = left + 1; right < paths.length; right += 1) {
      if (pathsOverlap(paths[left].path, paths[right].path)) {
        throw new Error(`Starter output paths overlap: ${paths[left].label} and ${paths[right].label}.`);
      }
    }
  }
}

export async function prepareTemplateStarterDeck(options) {
  const workspaceDir = path.resolve(options.workspace);
  const pptxPath = path.resolve(options.pptxPath);
  const mapPath = path.resolve(options.mapPath);
  const out = path.resolve(options.out);
  const previewDir = path.resolve(options.previewDir || path.join(workspaceDir, "template-starter-preview"));
  const layoutDir = path.resolve(options.layoutDir || path.join(workspaceDir, "template-starter-layout"));
  const inspectPath = path.resolve(options.inspectPath || path.join(workspaceDir, "template-inspect", "template-inspect.ndjson"));
  const contactSheetPath = options.contactSheetPath ? path.resolve(options.contactSheetPath) : undefined;
  const manifestPath = manifestPathFor(out);
  const scale = options.scale === undefined ? 1 : Number(options.scale);

  if (!Number.isFinite(scale) || scale <= 0) throw new Error("--scale must be a positive number");
  await fs.mkdir(workspaceDir, { recursive: true });
  for (const [label, filePath] of [["map", mapPath], ["inspect", inspectPath]]) {
    if (!isWithin(filePath, workspaceDir)) throw new Error(`${label} must be stored inside the template workspace: ${filePath}`);
  }
  const writePaths = [
    { label: "output PPTX", path: out },
    { label: "output manifest", path: manifestPath },
    { label: "preview directory", path: previewDir },
    { label: "layout directory", path: layoutDir },
    ...(contactSheetPath ? [{ label: "contact sheet", path: contactSheetPath }] : []),
  ];
  for (const { label, path: writePath } of writePaths) {
    if (!isWithin(writePath, workspaceDir) || writePath === workspaceDir) {
      throw new Error(`${label} must be a dedicated path inside the template workspace: ${writePath}`);
    }
  }
  assertDisjointOutputs(writePaths);
  if (pptxPath === out) throw new Error("Starter output must be distinct from the immutable source PPTX.");
  if (!out.toLowerCase().endsWith(".pptx")) throw new Error("Starter output must use a .pptx extension.");

  await Promise.all([
    assertRegularFile(pptxPath, "source PPTX"),
    assertRegularFile(mapPath, "template frame map", MAX_FRAME_MAP_BYTES),
    assertRegularFile(inspectPath, "template inspection"),
    ...writePaths.map(({ path: writePath, label }) => assertAbsent(writePath, label)),
  ]);
  const [sourceBytes, mapBytes, inspectBytes] = await Promise.all([
    fs.readFile(pptxPath),
    fs.readFile(mapPath),
    fs.readFile(inspectPath),
  ]);
  let map;
  try {
    map = JSON.parse(mapBytes.toString("utf8"));
  } catch (error) {
    throw new Error(`template frame map must be valid JSON: ${error.message}`);
  }

  const { FileBlob, PresentationFile } = await importArtifactTool(workspaceDir);
  let presentation = await PresentationFile.importPptx(fileBlob(FileBlob, sourceBytes, path.basename(pptxPath)));
  const sourceSlides = slidesFromPresentation(presentation);
  const sourceSlideCount = sourceSlides.length;
  const outputSlides = normalizeOutputSlides(map, sourceSlideCount);
  const planCheck = await validateTemplatePlan({
    workspace: workspaceDir,
    mapPath,
    inspectPath,
    sourceSlideCount,
    writeReport: false,
  });
  if (planCheck.status === "fail") {
    const summary = planCheck.issues.filter((item) => item.severity === "fail").slice(0, 8)
      .map((item) => `- ${item.id}: ${item.message}`).join("\n");
    throw new Error(["template-frame-map.json failed template plan validation.", summary].filter(Boolean).join("\n"));
  }

  const sourceRecords = inspectRecordsBySlide(presentation);
  const sourceVisuals = new Map();
  for (const sourceSlide of new Set(outputSlides.map((entry) => Number(entry.sourceSlide)))) {
    sourceVisuals.set(sourceSlide, await modelVisualSha256(sourceSlides[sourceSlide - 1]));
  }

  const cloneBoundaries = [];
  let workingBytes = sourceBytes;
  for (const entry of outputSlides) {
    const slides = slidesFromPresentation(presentation);
    const expectedBefore = sourceSlideCount + cloneBoundaries.length;
    if (slides.length !== expectedBefore) throw new Error(`Clone boundary expected ${expectedBefore} slides; found ${slides.length}.`);
    const sourceSlide = slides[Number(entry.sourceSlide) - 1];
    const capability = sourceSlide.cloneCapability;
    if (!capability?.known || !capability.supported) {
      throw new Error(`outputSlide ${entry.outputSlide} cannot clone sourceSlide ${entry.sourceSlide}: ${capability?.blockedReason || "clone capability is unavailable"}`);
    }
    const clone = sourceSlide.duplicate();
    clone.moveTo(presentation.slides.count - 1);
    workingBytes = await exportBytes(PresentationFile, presentation);
    presentation = await PresentationFile.importPptx(fileBlob(FileBlob, workingBytes, `template-clone-${entry.outputSlide}.pptx`));
    const reimportedSlides = slidesFromPresentation(presentation);
    if (reimportedSlides.length !== expectedBefore + 1) {
      throw new Error(`outputSlide ${entry.outputSlide} clone boundary did not add exactly one slide.`);
    }
    if (reimportedSlides.at(-1).name !== sourceSlide.name) {
      throw new Error(`outputSlide ${entry.outputSlide} clone changed its inherited slide name.`);
    }
    cloneBoundaries.push({
      outputSlide: Number(entry.outputSlide),
      sourceSlide: Number(entry.sourceSlide),
      clonedPartCount: Number(capability.clonedPartCount || 0),
      sharedPartCount: Number(capability.sharedPartCount || 0),
      exportReimported: true,
    });
  }

  const beforeDelete = slidesFromPresentation(presentation);
  const deletionCapabilities = beforeDelete.slice(0, sourceSlideCount).map((slide, index) => ({
    sourceSlide: index + 1,
    ...slide.deletionCapability,
  }));
  const blockedDeletion = deletionCapabilities.find((capability) => !capability.known || !capability.supported);
  if (blockedDeletion) {
    throw new Error(`Cannot remove sourceSlide ${blockedDeletion.sourceSlide} after cloning: ${blockedDeletion.blockedReason || "deletion capability is unavailable"}`);
  }
  for (let index = sourceSlideCount - 1; index >= 0; index -= 1) beforeDelete[index].delete();
  workingBytes = await exportBytes(PresentationFile, presentation);
  const starter = await PresentationFile.importPptx(fileBlob(FileBlob, workingBytes, path.basename(out)));
  const starterSlides = slidesFromPresentation(starter);
  if (starterSlides.length !== outputSlides.length) {
    throw new Error(`Starter reimport expected ${outputSlides.length} slides; found ${starterSlides.length}.`);
  }
  const verification = starter.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Starter presentation verification failed: ${verification.ndjson}`);

  const starterRecords = inspectRecordsBySlide(starter);
  const translatedSlides = [];
  for (let index = 0; index < outputSlides.length; index += 1) {
    const entry = outputSlides[index];
    const outputSlide = index + 1;
    const sourceSlide = Number(entry.sourceSlide);
    const visualSha256 = await modelVisualSha256(starterSlides[index]);
    if (visualSha256 !== sourceVisuals.get(sourceSlide)) {
      throw new Error(`outputSlide ${outputSlide} is not visually equivalent to sourceSlide ${sourceSlide}.`);
    }
    const translation = translateStarterTargets(
      sourceRecords.get(sourceSlide) || [],
      starterRecords.get(outputSlide) || [],
      entry,
    );
    translatedSlides.push({ entry, visualSha256, ...translation });
  }

  if (sha256(await fs.readFile(pptxPath)) !== sha256(sourceBytes)) throw new Error("Source PPTX changed during starter transaction.");
  if (sha256(await fs.readFile(mapPath)) !== sha256(mapBytes)) throw new Error("Template frame map changed during starter transaction.");
  if (sha256(await fs.readFile(inspectPath)) !== sha256(inspectBytes)) throw new Error("Template inspection changed during starter transaction.");

  const stagingDir = await fs.mkdtemp(path.join(workspaceDir, ".office-kit-template-starter-"));
  const stagedOut = path.join(stagingDir, "template-starter.pptx");
  const stagedManifest = path.join(stagingDir, "template-starter.manifest.json");
  const stagedPreviewDir = path.join(stagingDir, "preview");
  const stagedLayoutDir = path.join(stagingDir, "layout");
  const stagedContactSheet = contactSheetPath ? path.join(stagingDir, "contact-sheet.png") : undefined;
  const published = [];
  try {
    await fs.writeFile(stagedOut, workingBytes);
    await fs.mkdir(stagedPreviewDir, { recursive: true });
    await fs.mkdir(stagedLayoutDir, { recursive: true });
    const previewPaths = [];
    const layoutPaths = [];
    for (let index = 0; index < starterSlides.length; index += 1) {
      const padded = padSlideNumber(index + 1);
      const previewPath = path.join(stagedPreviewDir, `starter-slide-${padded}.png`);
      await saveBlobToFile(await starter.export({ slide: starterSlides[index], format: "png", scale }), previewPath);
      previewPaths.push(previewPath);
      const layoutPath = path.join(stagedLayoutDir, `starter-slide-${padded}.layout.json`);
      await saveBlobToFile(await starter.export({ slide: starterSlides[index], format: "layout" }), layoutPath);
      layoutPaths.push(layoutPath);
    }
    await runContactSheet(previewPaths, stagedContactSheet);

    const manifest = {
      schema: "office-kit.template-starter.v1",
      status: "succeeded",
      source: { path: pptxPath, sha256: sha256(sourceBytes), bytes: sourceBytes.length, immutable: true },
      frameMap: { path: mapPath, sha256: sha256(mapBytes), bytes: mapBytes.length },
      inspection: { path: inspectPath, sha256: sha256(inspectBytes), bytes: inspectBytes.length },
      output: { path: out, sha256: sha256(workingBytes), bytes: workingBytes.length },
      provider: { actual: "office-kit", silentFallback: false },
      savePolicy: { strategy: "rewrite", sourceMutation: false, overwrite: false },
      operation: {
        type: "source-bound-template-frame-map",
        sourceSlideCount,
        outputSlideCount: outputSlides.length,
        cloneBoundaries,
        sourceDeletionBoundary: { sourceSlideCount, exportReimported: true },
      },
      validation: {
        plan: { status: planCheck.status, issueCount: planCheck.issueCount },
        sourceImmutable: true,
        oneClonePerExportBoundary: true,
        finalExportReimported: true,
        exactOutputOrder: true,
        inheritedModelVisualsEquivalent: true,
        locatorTranslationComplete: true,
        verify: { ok: verification.ok },
      },
      previewDir,
      layoutDir,
      ...(contactSheetPath ? { contactSheet: contactSheetPath } : {}),
      slides: translatedSlides.map(({ entry, visualSha256, locators, editTargets }, index) => ({
        outputSlide: index + 1,
        sourceSlide: Number(entry.sourceSlide),
        narrativeRole: entry.narrativeRole,
        reuseMode: entry.reuseMode,
        visualSha256,
        editTargetCount: entry.editTargets.length,
        locators,
        editTargets,
        previewPath: path.join(previewDir, path.basename(previewPaths[index])),
        previewRelativePath: relativeFromWorkspace(workspaceDir, path.join(previewDir, path.basename(previewPaths[index]))),
        layoutPath: path.join(layoutDir, path.basename(layoutPaths[index])),
        layoutRelativePath: relativeFromWorkspace(workspaceDir, path.join(layoutDir, path.basename(layoutPaths[index]))),
      })),
    };
    await writeJson(stagedManifest, manifest);

    await publishDirectoryNoReplace(stagedPreviewDir, previewDir, "preview directory");
    published.push(previewDir);
    await publishDirectoryNoReplace(stagedLayoutDir, layoutDir, "layout directory");
    published.push(layoutDir);
    if (contactSheetPath) {
      await publishFileNoReplace(stagedContactSheet, contactSheetPath, "contact sheet");
      published.push(contactSheetPath);
    }
    await publishFileNoReplace(stagedOut, out, "output PPTX");
    published.push(out);
    await publishFileNoReplace(stagedManifest, manifestPath, "output manifest");
    published.push(manifestPath);
    return { outputPath: out, manifestPath, manifest };
  } catch (error) {
    await Promise.all(published.map((publishedPath) => fs.rm(publishedPath, { recursive: true, force: true })));
    throw error;
  } finally {
    await fs.rm(stagingDir, { recursive: true, force: true });
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log(usage());
    return;
  }
  const result = await prepareTemplateStarterDeck({
    workspace: requireArg(args, "workspace"),
    pptxPath: requireArg(args, "pptx"),
    mapPath: requireArg(args, "map"),
    out: requireArg(args, "out"),
    previewDir: args["preview-dir"],
    layoutDir: args["layout-dir"],
    inspectPath: args.inspect,
    contactSheetPath: args["contact-sheet"],
    scale: args.scale,
  });
  console.log(JSON.stringify({ outputPath: result.outputPath, manifestPath: result.manifestPath, outputSha256: result.manifest.output.sha256 }, null, 2));
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main().catch((error) => {
    console.error(error.stack || error.message || String(error));
    console.error(usage());
    process.exit(1);
  });
}
