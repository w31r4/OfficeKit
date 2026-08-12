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

const MAX_JSON_BYTES = 1_000_000;
const MAX_OPERATIONS = 512;
const MAX_ASSET_BYTES = 50_000_000;
const INSPECT_KINDS = "slide,textbox,shape,image,table,chart,connector,groupShape,nativeObject";
const MAX_INSPECT_CHARS = 20_000_000;

const ACTION_OPERATIONS = Object.freeze({
  keep: new Set(),
  rewrite: new Set(["set-text", "replace-text", "set-table-cell", "set-chart-title", "set-chart-series-values"]),
  "fill-placeholder": new Set(["set-text", "replace-text", "set-table-cell", "set-chart-title", "set-chart-series-values"]),
  "rewrite-and-reposition": new Set(["set-text", "replace-text", "set-table-cell", "set-chart-title", "set-chart-series-values", "set-position"]),
  replace: new Set(["replace-image"]),
});

function usage() {
  return [
    "Usage:",
    "  officekit run template_following_scripts/apply_template_edit_plan.mjs --workspace <dir> --starter <starter.pptx> --manifest <starter.manifest.json> --plan <template-edit-plan.json> --out <final.pptx> [options]",
    "",
    "Options:",
    "  --audit <path>          Final audit JSON. Defaults beside --out.",
    "  --preview-dir <dir>    Final slide PNGs. Defaults to <workspace>/template-final-preview.",
    "  --layout-dir <dir>     Final layout JSON. Defaults to <workspace>/template-final-layout.",
    "  --contact-sheet <path> Optional PNG contact sheet path.",
    "  --scale <n>            Render scale. Defaults to 1.",
    "",
    "Applies one hash-bound, typed edit plan to a generated template starter.",
    "The command validates every mapped target, fails closed on stale or unsupported",
    "operations, reimports and verifies the result, then publishes with no overwrite.",
  ].join("\n");
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value;
}

function requiredSha256(value, label) {
  const normalized = requiredText(value, label).toLowerCase();
  if (!/^[0-9a-f]{64}$/.test(normalized)) throw new TypeError(`${label} must be a lowercase SHA-256 hex digest.`);
  return normalized;
}

function exactJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function exactPosition(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  const result = Object.fromEntries(["left", "top", "width", "height"].map((key) => [key, Number(value[key])]));
  if (Object.values(result).some((item) => !Number.isFinite(item)) || result.width <= 0 || result.height <= 0) {
    throw new TypeError(`${label} must contain finite left/top and positive width/height.`);
  }
  return result;
}

function operationKey(operation) {
  if (operation.type === "set-table-cell") return `${operation.type}:${operation.row}:${operation.column}`;
  if (operation.type === "set-chart-series-values") return `${operation.type}:${operation.seriesIndex}`;
  return operation.type;
}

function countSubstring(value, needle) {
  let count = 0;
  let offset = 0;
  while ((offset = value.indexOf(needle, offset)) >= 0) {
    count += 1;
    offset += needle.length;
  }
  return count;
}

function imageBytes(dataUrl, label) {
  const match = /^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/=\r\n]+)$/i.exec(String(dataUrl || ""));
  if (!match) throw new Error(`${label} must be an embedded PNG or JPEG image.`);
  return { contentType: match[1].toLowerCase(), bytes: Buffer.from(match[2], "base64") };
}

function imageAsset(bytes, label) {
  if (bytes.subarray(0, 8).toString("hex") === "89504e470d0a1a0a") return { contentType: "image/png", bytes };
  if (bytes.subarray(0, 3).toString("hex") === "ffd8ff") return { contentType: "image/jpeg", bytes };
  throw new Error(`${label} must be a PNG or JPEG image.`);
}

function inspectRecordsBySlide(presentation) {
  const inspection = presentation.inspect({ kind: INSPECT_KINDS, maxChars: MAX_INSPECT_CHARS });
  if (inspection.truncated) throw new Error(`Template edit locator inspection exceeded ${MAX_INSPECT_CHARS} characters.`);
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

function identityShape(record) {
  return JSON.stringify({
    kind: record.kind,
    name: record.name || "",
    nativeId: record.nativeId ?? null,
    creationId: record.creationId ?? null,
    nativeKind: record.nativeKind || "",
  });
}

function manifestTargets(manifest) {
  if (manifest?.schema !== "office-kit.template-starter.v1" || manifest.status !== "succeeded") {
    throw new Error("Starter manifest must be a successful office-kit.template-starter.v1 record.");
  }
  requiredSha256(manifest.output?.sha256, "manifest.output.sha256");
  if (!Array.isArray(manifest.slides) || manifest.slides.length === 0) throw new Error("Starter manifest must contain slides.");
  const targets = new Map();
  const authorizedElementIds = new Map();
  for (let slideIndex = 0; slideIndex < manifest.slides.length; slideIndex += 1) {
    const slide = manifest.slides[slideIndex];
    if (Number(slide.outputSlide) !== slideIndex + 1 || !Array.isArray(slide.editTargets)) {
      throw new Error(`Starter manifest slide ${slideIndex + 1} is malformed.`);
    }
    for (const target of slide.editTargets) {
      const targetIndex = Number(target.targetIndex);
      const action = String(target.action || "");
      if (!Number.isInteger(targetIndex) || targetIndex < 0 || !Array.isArray(target.starterElementIds) || target.starterElementIds.length === 0) {
        throw new Error(`Starter manifest outputSlide ${slide.outputSlide} has a malformed edit target.`);
      }
      const key = `${slide.outputSlide}:${targetIndex}`;
      if (targets.has(key)) throw new Error(`Starter manifest repeats edit target ${key}.`);
      for (const starterElementId of target.starterElementIds) {
        if (typeof starterElementId !== "string" || !starterElementId) throw new Error(`Starter manifest edit target ${key} contains an invalid element ID.`);
        const previous = authorizedElementIds.get(starterElementId);
        if (previous) throw new Error(`Starter manifest element ${starterElementId} is ambiguously authorized by targets ${previous} and ${key}.`);
        authorizedElementIds.set(starterElementId, key);
      }
      targets.set(key, { outputSlide: Number(slide.outputSlide), targetIndex, action, starterElementIds: [...target.starterElementIds] });
    }
  }
  return targets;
}

function normalizePlan(plan, manifest, manifestBytes, starterBytes) {
  if (plan?.schema !== "office-kit.template-edit-plan.v1") throw new Error("Edit plan schema must be office-kit.template-edit-plan.v1.");
  if (requiredSha256(plan.starterSha256, "plan.starterSha256") !== sha256(starterBytes)) {
    throw new Error("Edit plan starterSha256 does not match the starter PPTX.");
  }
  if (requiredSha256(plan.manifestSha256, "plan.manifestSha256") !== sha256(manifestBytes)) {
    throw new Error("Edit plan manifestSha256 does not match the starter manifest.");
  }
  if (!Array.isArray(plan.targets)) throw new Error("Edit plan targets must be an array.");
  const expected = manifestTargets(manifest);
  if (plan.targets.length !== expected.size) {
    throw new Error(`Edit plan must cover all ${expected.size} starter edit targets exactly once; found ${plan.targets.length}.`);
  }
  let operationCount = 0;
  const normalized = [];
  const seen = new Set();
  for (const entry of plan.targets) {
    const outputSlide = Number(entry?.outputSlide);
    const targetIndex = Number(entry?.targetIndex);
    const key = `${outputSlide}:${targetIndex}`;
    if (!Number.isInteger(outputSlide) || outputSlide < 1 || !Number.isInteger(targetIndex) || targetIndex < 0 || seen.has(key)) {
      throw new Error(`Edit plan target ${key} is invalid or duplicated.`);
    }
    seen.add(key);
    const target = expected.get(key);
    if (!target) throw new Error(`Edit plan target ${key} is not authorized by the starter manifest.`);
    if (target.action === "delete") throw new Error(`Edit target ${key} requests delete, which this bounded transaction does not support.`);
    const allowed = ACTION_OPERATIONS[target.action];
    if (!allowed) throw new Error(`Edit target ${key} uses unsupported frame-map action ${JSON.stringify(target.action)}.`);
    if (!Array.isArray(entry.operations)) throw new Error(`Edit plan target ${key} operations must be an array.`);
    if (target.action === "keep" && entry.operations.length !== 0) throw new Error(`Keep target ${key} cannot contain mutations.`);
    if (target.action !== "keep" && entry.operations.length === 0) throw new Error(`Edit target ${key} requires at least one typed operation.`);
    const coveredElements = new Set();
    const uniqueOperations = new Set();
    const operations = entry.operations.map((operation, index) => {
      const type = String(operation?.type || "");
      if (!allowed.has(type)) throw new Error(`Edit target ${key} operation ${index} cannot use ${JSON.stringify(type)} for action ${target.action}.`);
      const elementIndex = operation.elementIndex === undefined ? (target.starterElementIds.length === 1 ? 0 : NaN) : Number(operation.elementIndex);
      if (!Number.isInteger(elementIndex) || elementIndex < 0 || elementIndex >= target.starterElementIds.length) {
        throw new Error(`Edit target ${key} operation ${index} needs a valid elementIndex for ${target.starterElementIds.length} element(s).`);
      }
      const unique = `${elementIndex}:${operationKey(operation)}`;
      if (uniqueOperations.has(unique)) throw new Error(`Edit target ${key} repeats operation ${unique}.`);
      uniqueOperations.add(unique);
      coveredElements.add(elementIndex);
      operationCount += 1;
      return { ...operation, type, elementIndex };
    });
    if (target.action !== "keep" && coveredElements.size !== target.starterElementIds.length) {
      throw new Error(`Edit target ${key} must mutate every mapped starter element.`);
    }
    if (target.action === "rewrite-and-reposition" &&
        (!operations.some((operation) => operation.type === "set-position") || !operations.some((operation) => operation.type !== "set-position"))) {
      throw new Error(`Edit target ${key} action rewrite-and-reposition requires set-position plus a content mutation.`);
    }
    normalized.push({ ...target, operations });
  }
  if (operationCount === 0) throw new Error("Edit plan does not contain a mutation; deliver the verified starter instead.");
  if (operationCount > MAX_OPERATIONS) throw new Error(`Edit plan exceeds the ${MAX_OPERATIONS}-operation budget.`);
  return { targets: normalized, operationCount };
}

function assertTargetSlide(target, outputSlide, label) {
  if (Number(target?.slide?.index) + 1 !== outputSlide) throw new Error(`${label} resolved outside outputSlide ${outputSlide}.`);
}

function operationHandlers(context) {
  return {
    "set-text": async (target, operation, label) => {
      if (!Object.hasOwn(operation, "expectedText") || typeof operation.expectedText !== "string") {
        throw new Error(`${label}.expectedText must be a string, including an explicit empty string when applicable.`);
      }
      const text = requiredText(operation.text, `${label}.text`);
      if (typeof target?.text?.set !== "function") throw new Error(`${label} requires a text-bearing shape.`);
      if (target.text.value !== operation.expectedText) throw new Error(`${label} text precondition failed.`);
      target.text.set(text);
      return { before: operation.expectedText, after: text, verify: (roundTrip) => roundTrip?.text?.value === text };
    },
    "replace-text": async (target, operation, label) => {
      const expected = requiredText(operation.expectedText, `${label}.expectedText`);
      const replacement = requiredText(operation.text, `${label}.text`);
      const before = String(target?.text?.value ?? "");
      if (typeof target?.text?.replace !== "function") throw new Error(`${label} requires a text-bearing shape.`);
      if (countSubstring(before, expected) !== 1) throw new Error(`${label} expectedText must occur exactly once; found ${countSubstring(before, expected)}.`);
      target.text.replace(expected, replacement);
      const after = before.replace(expected, replacement);
      if (target.text.value !== after) throw new Error(`${label} text replacement crossed an unsupported run boundary.`);
      return { before: expected, after: replacement, verify: (roundTrip) => roundTrip?.text?.value === after };
    },
    "set-position": async (target, operation, label) => {
      const expected = exactPosition(operation.expectedPosition, `${label}.expectedPosition`);
      const value = exactPosition(operation.position, `${label}.position`);
      if (!exactJson(target?.position, expected)) throw new Error(`${label} position precondition failed.`);
      target.position = value;
      return { before: expected, after: value, verify: (roundTrip) => exactJson(roundTrip?.position, value) };
    },
    "set-table-cell": async (target, operation, label) => {
      const row = Number(operation.row);
      const column = Number(operation.column);
      if (!Number.isInteger(row) || !Number.isInteger(column) || typeof target?.getCell !== "function") {
        throw new Error(`${label} requires integer row/column and a table target.`);
      }
      if (!Object.hasOwn(operation, "expectedValue") || !Object.hasOwn(operation, "value")) {
        throw new Error(`${label} requires expectedValue and value.`);
      }
      const cell = target.getCell(row, column);
      if (!exactJson(cell.value, operation.expectedValue)) throw new Error(`${label} table-cell precondition failed.`);
      cell.value = operation.value;
      return { row, column, before: operation.expectedValue, after: operation.value, verify: (roundTrip) => exactJson(roundTrip?.getCell(row, column).value, operation.value) };
    },
    "set-chart-title": async (target, operation, label) => {
      const expected = requiredText(operation.expectedTitle, `${label}.expectedTitle`);
      const title = requiredText(operation.title, `${label}.title`);
      if (!Array.isArray(target?.series)) throw new Error(`${label} requires a chart target.`);
      if (target.title !== expected) throw new Error(`${label} chart-title precondition failed.`);
      target.title = title;
      return { before: expected, after: title, verify: (roundTrip) => roundTrip?.title === title };
    },
    "set-chart-series-values": async (target, operation, label) => {
      const seriesIndex = Number(operation.seriesIndex);
      if (!Number.isInteger(seriesIndex) || seriesIndex < 0 || !Array.isArray(target?.series?.[seriesIndex]?.values)) {
        throw new Error(`${label} requires an existing chart seriesIndex.`);
      }
      const expected = operation.expectedValues;
      const values = operation.values;
      if (!Array.isArray(expected) || !Array.isArray(values) || values.length !== expected.length || values.some((value) => !Number.isFinite(Number(value)))) {
        throw new Error(`${label} requires same-length expectedValues/values with finite replacement values.`);
      }
      if (!exactJson(target.series[seriesIndex].values, expected)) throw new Error(`${label} chart-series precondition failed.`);
      target.series[seriesIndex].values = [...values];
      return { seriesIndex, before: expected, after: values, verify: (roundTrip) => exactJson(roundTrip?.series?.[seriesIndex]?.values, values) };
    },
    "replace-image": async (target, operation, label) => {
      if (typeof target?.replace !== "function" || typeof target?.dataUrl !== "string") throw new Error(`${label} requires an embedded image target.`);
      const current = imageBytes(target.dataUrl, `${label} source image`);
      const expectedSourceSha256 = requiredSha256(operation.expectedSourceSha256, `${label}.expectedSourceSha256`);
      if (sha256(current.bytes) !== expectedSourceSha256) throw new Error(`${label} source-image precondition failed.`);
      const assetPath = path.resolve(context.planDir, requiredText(operation.assetPath, `${label}.assetPath`));
      if (!isWithin(assetPath, context.workspaceDir)) throw new Error(`${label} assetPath must be inside the template workspace.`);
      await assertRegularFile(assetPath, `${label} asset`, MAX_ASSET_BYTES);
      const bytes = await fs.readFile(assetPath);
      const expectedAssetSha256 = requiredSha256(operation.assetSha256, `${label}.assetSha256`);
      if (sha256(bytes) !== expectedAssetSha256) throw new Error(`${label} assetSha256 does not match ${assetPath}.`);
      const asset = imageAsset(bytes, `${label} asset`);
      target.replace({ dataUrl: `data:${asset.contentType};base64,${bytes.toString("base64")}`, uri: undefined, prompt: undefined });
      context.assets.set(assetPath, { bytes, sha256: expectedAssetSha256 });
      return {
        assetPath,
        beforeSha256: expectedSourceSha256,
        afterSha256: expectedAssetSha256,
        verify: (roundTrip) => sha256(imageBytes(roundTrip?.dataUrl, `${label} round-trip image`).bytes) === expectedAssetSha256,
      };
    },
  };
}

function auditPathFor(out) {
  return out.toLowerCase().endsWith(".pptx") ? `${out.slice(0, -5)}.audit.json` : `${out}.audit.json`;
}

function assertDisjointOutputs(paths) {
  for (let left = 0; left < paths.length; left += 1) {
    for (let right = left + 1; right < paths.length; right += 1) {
      if (pathsOverlap(paths[left].path, paths[right].path)) {
        throw new Error(`Template edit output paths overlap: ${paths[left].label} and ${paths[right].label}.`);
      }
    }
  }
}

async function assertSnapshotsUnchanged(inputs) {
  for (const [inputPath, snapshot, label] of inputs) {
    if (sha256(await fs.readFile(inputPath)) !== sha256(snapshot)) {
      throw new Error(`${label} changed during the edit transaction.`);
    }
  }
}

export async function applyTemplateEditPlan(options) {
  const workspaceDir = path.resolve(options.workspace);
  const starterPath = path.resolve(options.starterPath);
  const manifestPath = path.resolve(options.manifestPath);
  const planPath = path.resolve(options.planPath);
  const out = path.resolve(options.out);
  const auditPath = path.resolve(options.auditPath || auditPathFor(out));
  const previewDir = path.resolve(options.previewDir || path.join(workspaceDir, "template-final-preview"));
  const layoutDir = path.resolve(options.layoutDir || path.join(workspaceDir, "template-final-layout"));
  const contactSheetPath = options.contactSheetPath ? path.resolve(options.contactSheetPath) : undefined;
  const scale = options.scale === undefined ? 1 : Number(options.scale);

  if (!Number.isFinite(scale) || scale <= 0) throw new Error("--scale must be a positive number");
  if (!out.toLowerCase().endsWith(".pptx")) throw new Error("Template edit output must use a .pptx extension.");
  if (starterPath === out) throw new Error("Template edit output must be distinct from the immutable starter PPTX.");
  await fs.mkdir(workspaceDir, { recursive: true });
  for (const [label, inputPath] of [["starter", starterPath], ["manifest", manifestPath], ["plan", planPath]]) {
    if (!isWithin(inputPath, workspaceDir)) throw new Error(`${label} must be stored inside the template workspace: ${inputPath}`);
  }
  for (const [label, directory] of [["preview directory", previewDir], ["layout directory", layoutDir]]) {
    if (!isWithin(directory, workspaceDir) || directory === workspaceDir) throw new Error(`${label} must be a dedicated path inside the template workspace.`);
  }
  if (contactSheetPath && (!isWithin(contactSheetPath, workspaceDir) || contactSheetPath === workspaceDir)) {
    throw new Error("contact sheet must be stored inside the template workspace.");
  }
  const writePaths = [
    { label: "output PPTX", path: out },
    { label: "output audit", path: auditPath },
    { label: "preview directory", path: previewDir },
    { label: "layout directory", path: layoutDir },
    ...(contactSheetPath ? [{ label: "contact sheet", path: contactSheetPath }] : []),
  ];
  assertDisjointOutputs(writePaths);
  await Promise.all([
    assertRegularFile(starterPath, "starter PPTX"),
    assertRegularFile(manifestPath, "starter manifest", MAX_JSON_BYTES),
    assertRegularFile(planPath, "template edit plan", MAX_JSON_BYTES),
    ...writePaths.map(({ path: writePath, label }) => assertAbsent(writePath, label)),
  ]);
  const [starterBytes, manifestBytes, planBytes] = await Promise.all([
    fs.readFile(starterPath),
    fs.readFile(manifestPath),
    fs.readFile(planPath),
  ]);
  let manifest;
  let plan;
  try {
    manifest = JSON.parse(manifestBytes.toString("utf8"));
    plan = JSON.parse(planBytes.toString("utf8"));
  } catch (error) {
    throw new Error(`Template edit inputs must be valid JSON: ${error.message}`);
  }
  if (manifest.output.sha256 !== sha256(starterBytes)) throw new Error("Starter manifest output hash does not match the starter PPTX.");
  const normalizedPlan = normalizePlan(plan, manifest, manifestBytes, starterBytes);

  const { FileBlob, PresentationFile } = await importArtifactTool(workspaceDir);
  const presentation = await PresentationFile.importPptx(fileBlob(FileBlob, starterBytes, path.basename(starterPath)));
  const slides = slidesFromPresentation(presentation);
  if (slides.length !== manifest.slides.length) throw new Error("Starter slide count does not match its manifest.");
  const beforeRecords = inspectRecordsBySlide(presentation);
  const untouchedSlides = new Set(slides.map((_, index) => index + 1));
  const untouchedVisuals = new Map(await Promise.all(slides.map(async (slide, index) => [index + 1, await modelVisualSha256(slide)])));
  const assets = new Map();
  const handlers = operationHandlers({ workspaceDir, planDir: path.dirname(planPath), assets });
  const verifications = [];
  const auditOperations = [];

  for (const targetPlan of normalizedPlan.targets) {
    const key = `${targetPlan.outputSlide}:${targetPlan.targetIndex}`;
    const records = beforeRecords.get(targetPlan.outputSlide) || [];
    for (const operation of targetPlan.operations) {
      const starterElementId = targetPlan.starterElementIds[operation.elementIndex];
      const target = presentation.resolve(starterElementId);
      if (!target) throw new Error(`Edit target ${key} element ${operation.elementIndex} no longer resolves.`);
      assertTargetSlide(target, targetPlan.outputSlide, `Edit target ${key}`);
      const recordIndex = records.findIndex((record) => record.id === starterElementId);
      if (recordIndex < 0) throw new Error(`Edit target ${key} element ${operation.elementIndex} is absent from bounded inspection.`);
      const label = `Edit target ${key} ${operation.type}`;
      const outcome = await handlers[operation.type](target, operation, label);
      untouchedSlides.delete(targetPlan.outputSlide);
      const auditIndex = auditOperations.length;
      verifications.push({
        label,
        outputSlide: targetPlan.outputSlide,
        sourceRecordIndex: recordIndex,
        sourceIdentity: identityShape(records[recordIndex]),
        verify: outcome.verify,
        auditIndex,
      });
      auditOperations.push({
        outputSlide: targetPlan.outputSlide,
        targetIndex: targetPlan.targetIndex,
        elementIndex: operation.elementIndex,
        starterElementId,
        action: targetPlan.action,
        type: operation.type,
        executed: true,
        ...Object.fromEntries(Object.entries(outcome).filter(([name]) => name !== "verify" && name !== "bytes")),
      });
    }
  }

  const outputBytes = await exportBytes(PresentationFile, presentation);
  const roundTrip = await PresentationFile.importPptx(fileBlob(FileBlob, outputBytes, path.basename(out)));
  const roundTripSlides = slidesFromPresentation(roundTrip);
  if (roundTripSlides.length !== slides.length) throw new Error("Template edit changed slide topology.");
  const afterRecords = inspectRecordsBySlide(roundTrip);
  for (const verification of verifications) {
    const records = afterRecords.get(verification.outputSlide) || [];
    if (records.length !== (beforeRecords.get(verification.outputSlide) || []).length) {
      throw new Error(`${verification.label} changed the inherited element topology.`);
    }
    const record = records[verification.sourceRecordIndex];
    if (!record || identityShape(record) !== verification.sourceIdentity) {
      throw new Error(`${verification.label} changed inherited element identity or ordering.`);
    }
    const target = roundTrip.resolve(record.id);
    if (!target || !verification.verify(target)) throw new Error(`${verification.label} failed round-trip verification.`);
    auditOperations[verification.auditIndex].finalElementId = record.id;
  }
  const untouchedChecks = [];
  for (const outputSlide of untouchedSlides) {
    const before = untouchedVisuals.get(outputSlide);
    const after = await modelVisualSha256(roundTripSlides[outputSlide - 1]);
    if (before !== after) throw new Error(`Untouched outputSlide ${outputSlide} changed visually.`);
    untouchedChecks.push({ outputSlide, modelVisualSha256: after });
  }
  const verification = roundTrip.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Template edit verification failed: ${verification.ndjson}`);

  const immutableInputs = [
    [starterPath, starterBytes, "Starter PPTX"],
    [manifestPath, manifestBytes, "Starter manifest"],
    [planPath, planBytes, "Template edit plan"],
    ...[...assets].map(([assetPath, snapshot]) => [assetPath, snapshot.bytes, `Template edit asset ${assetPath}`]),
  ];
  await assertSnapshotsUnchanged(immutableInputs);

  const stagingDir = await fs.mkdtemp(path.join(workspaceDir, ".office-kit-template-edit-"));
  const stagedOut = path.join(stagingDir, "template-final.pptx");
  const stagedAudit = path.join(stagingDir, "template-final.audit.json");
  const stagedPreviewDir = path.join(stagingDir, "preview");
  const stagedLayoutDir = path.join(stagingDir, "layout");
  const stagedContactSheet = contactSheetPath ? path.join(stagingDir, "contact-sheet.png") : undefined;
  const published = [];
  try {
    await fs.writeFile(stagedOut, outputBytes);
    await fs.mkdir(stagedPreviewDir, { recursive: true });
    await fs.mkdir(stagedLayoutDir, { recursive: true });
    const previewPaths = [];
    const layoutPaths = [];
    for (let index = 0; index < roundTripSlides.length; index += 1) {
      const padded = padSlideNumber(index + 1);
      const previewPath = path.join(stagedPreviewDir, `final-slide-${padded}.png`);
      await saveBlobToFile(await roundTrip.export({ slide: roundTripSlides[index], format: "png", scale }), previewPath);
      previewPaths.push(previewPath);
      const layoutPath = path.join(stagedLayoutDir, `final-slide-${padded}.layout.json`);
      await saveBlobToFile(await roundTrip.export({ slide: roundTripSlides[index], format: "layout" }), layoutPath);
      layoutPaths.push(layoutPath);
    }
    await runContactSheet(previewPaths, stagedContactSheet);
    await assertSnapshotsUnchanged(immutableInputs);
    const audit = {
      schema: "office-kit.template-edit-audit.v1",
      status: "succeeded",
      provider: { actual: "office-kit", silentFallback: false },
      savePolicy: { strategy: "rewrite", sourceMutation: false, overwrite: false },
      source: { path: starterPath, sha256: sha256(starterBytes), bytes: starterBytes.length, immutable: true },
      starterManifest: { path: manifestPath, sha256: sha256(manifestBytes), schema: manifest.schema },
      editPlan: { path: planPath, sha256: sha256(planBytes), schema: plan.schema },
      assets: [...assets].map(([assetPath, snapshot]) => ({
        path: assetPath,
        relativePath: relativeFromWorkspace(workspaceDir, assetPath),
        sha256: snapshot.sha256,
        bytes: snapshot.bytes.length,
      })),
      output: { path: out, sha256: sha256(outputBytes), bytes: outputBytes.length },
      operation: { type: "source-bound-template-edit-plan", count: normalizedPlan.operationCount, operations: auditOperations },
      validation: {
        sourceImmutable: true,
        manifestImmutable: true,
        planImmutable: true,
        assetsImmutable: true,
        immutableInputsRecheckedBeforePublication: true,
        allMappedTargetsCovered: true,
        typedOperationsOnly: true,
        noSlideTopologyChange: true,
        locatorTranslationComplete: true,
        finalExportReimported: true,
        untouchedSlideVisualsEquivalent: true,
        untouchedSlides: untouchedChecks,
        verify: { ok: verification.ok },
      },
      previewDir,
      layoutDir,
      ...(contactSheetPath ? { contactSheet: contactSheetPath } : {}),
      slides: roundTripSlides.map((_, index) => ({
        outputSlide: index + 1,
        previewPath: path.join(previewDir, path.basename(previewPaths[index])),
        previewRelativePath: relativeFromWorkspace(workspaceDir, path.join(previewDir, path.basename(previewPaths[index]))),
        layoutPath: path.join(layoutDir, path.basename(layoutPaths[index])),
        layoutRelativePath: relativeFromWorkspace(workspaceDir, path.join(layoutDir, path.basename(layoutPaths[index]))),
      })),
    };
    await writeJson(stagedAudit, audit);

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
    await publishFileNoReplace(stagedAudit, auditPath, "output audit");
    published.push(auditPath);
    return { outputPath: out, auditPath, audit };
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
  const result = await applyTemplateEditPlan({
    workspace: requireArg(args, "workspace"),
    starterPath: requireArg(args, "starter"),
    manifestPath: requireArg(args, "manifest"),
    planPath: requireArg(args, "plan"),
    out: requireArg(args, "out"),
    auditPath: args.audit,
    previewDir: args["preview-dir"],
    layoutDir: args["layout-dir"],
    contactSheetPath: args["contact-sheet"],
    scale: args.scale,
  });
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256 }, null, 2));
}

const entry = process.argv[1] ? fileURLToPath(import.meta.url) === path.resolve(process.argv[1]) : false;
if (entry) main().catch((error) => {
  console.error(error?.stack || error);
  process.exit(1);
});
