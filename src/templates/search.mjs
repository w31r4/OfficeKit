import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const MODULE_DIRECTORY = path.dirname(fileURLToPath(import.meta.url));
const PACKAGE_ROOT = path.resolve(MODULE_DIRECTORY, "../..");
const MANAGED_SKILLS_MANIFEST = ".office-kit/skills.json";
const SIDECAR_NAME = "artifact-template.json";
const TEMPLATE_NAME_PATTERN = /^artifact-template-[a-z0-9]+(?:-[a-z0-9]+)*$/u;
const HASH_PATTERN = /^[a-f0-9]{64}$/u;
const MAX_SIDECAR_BYTES = 128 * 1024;
const MAX_SKILL_BYTES = 256 * 1024;
const MIN_PRESENTATION_EXAMPLES = 4;
const MAX_PRESENTATION_EXAMPLES = 6;
const DEFAULT_MAX_CANDIDATES = 5;
const MAX_CANDIDATES = 20;
const MAX_INTENT_VALUES = 20;
const MAX_REMOTE_REFERENCE_BYTES = 256 * 1024 * 1024;
const REMOTE_REFERENCE_HOSTS = new Set(["raw.githubusercontent.com"]);
const MIN_FIELD_MATCH = 0.45;
const AVOID_CONFLICT_MATCH = 0.72;
const BM25_K1 = 1.2;
const BM25_B = 0.75;
const VALID_KINDS = new Set(["document", "presentation", "spreadsheet"]);
const REFERENCE_EXTENSIONS = new Map([
  ["document", ".docx"],
  ["presentation", ".pptx"],
  ["spreadsheet", ".xlsx"],
]);
const VALID_DENSITIES = new Set(["sparse", "medium", "dense", "mixed"]);
const VALID_COLOR_MODES = new Set(["light", "dark", "neutral", "mixed"]);
const VALID_COMMITMENTS = new Set(["neutral", "opinionated"]);
const VALID_EDIT_LEVELS = new Set(["copy-only", "bounded-edit", "composable"]);
const VALID_PRESENTATION_EXAMPLE_ROLES = new Set([
  "cover",
  "section",
  "analysis",
  "data",
  "process",
  "comparison",
  "closing",
  "mixed",
]);
const VALID_IMAGE_SLOT_ROLES = new Set([
  "hero",
  "thumbnail",
  "avatar",
  "background",
  "logo",
  "diagram",
  "screenshot",
  "photo",
  "icon",
  "chart-source",
  "any",
]);
const VALID_IMAGE_SLOT_FITS = new Set(["contain", "cover", "stretch"]);
const VALID_IMAGE_SLOT_MASKS = new Set(["none", "rect", "roundRect", "ellipse", "custom"]);
const VALID_IMAGE_SLOT_RIGHTS = new Set([
  "user-provided",
  "generated",
  "permission",
  "public-domain",
  "cc0",
  "cc-by",
  "official-press-kit",
  "internal",
  "other",
]);
const IMAGE_SLOT_ID_PATTERN = /^[a-z][a-z0-9-]{0,63}$/u;
const BM25_FIELD_WEIGHTS = Object.freeze({
  identity: 1.5,
  useWhen: 4,
  audiences: 2,
  contentShapes: 2,
  tone: 1.25,
  structure: 1,
  density: 0.75,
  colorMode: 0.75,
});
export const TEMPLATE_SEARCH_USAGE = [
  "Usage: officekit template search --kind <document|spreadsheet|presentation>",
  "  [--purpose <phrase>]... [--audience <phrase>]...",
  "  [--content-shape <phrase>]... [--tone <phrase>]...",
  "  [--structure <phrase>]... [--density <value>] [--color-mode <value>]",
  "  [--operation <verified-operation>]... [--brand-sensitive]",
  "  [--tag <legacy-tag>]... [--id <artifact-template-id>]",
  "  [--root <absolute-template-root>]... [--max <1-20>] [--json]",
  "  officekit template fetch <artifact-template-id> [--cache-root <absolute-dir>] [--json]",
].join("\n");

export async function queryTemplates({
  kind,
  intent = null,
  tags = [],
  id = null,
  roots = null,
  maxCandidates = DEFAULT_MAX_CANDIDATES,
  projectPath = process.cwd(),
} = {}) {
  assertKind(kind);
  assertTemplateId(id, "--id", true);
  const normalizedTags = normalizeQueryTags(tags);
  const normalizedIntent = normalizeIntent(intent);
  if (!Number.isSafeInteger(maxCandidates) || maxCandidates < 1 || maxCandidates > MAX_CANDIDATES) {
    throw new Error(`maxCandidates must be an integer from 1 to ${MAX_CANDIDATES}.`);
  }

  const rootEntries = await resolveRoots(roots, projectPath);
  const discoveredCandidates = [];
  const candidates = [];
  const rejected = [];
  const invalid = [];
  const seenTemplatePaths = new Set();
  const claimedTemplateIds = new Set();

  for (const rootEntry of rootEntries) {
    for (const entry of await fs.readdir(rootEntry.path, { withFileTypes: true })) {
      if (!entry.isDirectory() || !TEMPLATE_NAME_PATTERN.test(entry.name)) continue;
      if (id != null && entry.name !== id) continue;
      if (claimedTemplateIds.has(entry.name)) continue;
      claimedTemplateIds.add(entry.name);
      const templatePath = path.join(rootEntry.path, entry.name);
      try {
        const canonicalTemplatePath = await fs.realpath(templatePath);
        if (seenTemplatePaths.has(canonicalTemplatePath)) continue;
        seenTemplatePaths.add(canonicalTemplatePath);
        const candidate = await readTemplate({
          expectedId: entry.name,
          root: rootEntry,
          templatePath,
        });
        if (candidate.kind !== kind) continue;
        if (id != null && candidate.id !== id) continue;
        discoveredCandidates.push(candidate);
      } catch (error) {
        invalid.push({
          id: entry.name,
          root: rootEntry.path,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }
  }

  invalid.sort((left, right) => left.id.localeCompare(right.id) || left.root.localeCompare(right.root));
  if (id != null) {
    const invalidRequested = invalid.find((entry) => entry.id === id);
    if (invalidRequested != null) {
      throw new Error(`Requested template ${id} is invalid: ${invalidRequested.error}`);
    }
    if (discoveredCandidates.length === 0) {
      throw new Error(`Requested template ${id} was not found for kind ${kind}.`);
    }
  }

  const bm25 = createBm25Context(
    discoveredCandidates,
    normalizedIntent,
    normalizedTags,
  );
  const assessments = discoveredCandidates.map((candidate) => ({
    candidate,
    assessment: assessCandidate(
      candidate,
      normalizedIntent,
      normalizedTags,
      bm25,
    ),
  }));
  const maximumBm25 = Math.max(
    0,
    ...assessments.map(({ assessment }) => assessment.match.bm25),
  );
  for (const { candidate, assessment } of assessments) {
    assessment.match.score =
      maximumBm25 === 0
        ? 0
        : roundScore((assessment.match.bm25 / maximumBm25) * 100);
    candidate.matchedTags = assessment.matchedTags;
    candidate.match = assessment.match;
    candidate.reviewFlags = assessment.reviewFlags;
    if (id == null && assessment.rejectionReasons.length > 0) {
      rejected.push({
        id: candidate.id,
        displayName: candidate.displayName,
        score: assessment.match.score,
        bm25: assessment.match.bm25,
        reasons: assessment.rejectionReasons,
        conflicts: assessment.match.conflicts,
        missingOperations: assessment.match.missingOperations,
      });
    } else {
      candidates.push(candidate);
    }
  }

  rejected.sort((left, right) =>
    right.score - left.score ||
    left.id.localeCompare(right.id)
  );
  candidates.sort((left, right) =>
    right.match.score - left.match.score ||
    right.matchedTags.length - left.matchedTags.length ||
    left.id.localeCompare(right.id) ||
    left.templateRoot.localeCompare(right.templateRoot)
  );

  return {
    schemaVersion: 2,
    kind,
    requestedId: id,
    queryIntent: normalizedIntent,
    queryTags: normalizedTags,
    ranking: {
      algorithm: "bm25f",
      k1: BM25_K1,
      b: BM25_B,
      queryTerms: bm25.queryTerms,
    },
    searchedRoots: rootEntries.map((entry) => ({
      path: entry.path,
      source: entry.source,
    })),
    candidates: candidates.slice(0, id == null ? maxCandidates : 1),
    rejected,
    invalid,
    retrievalStatus: candidates.length === 0 ? "none" : "candidates",
    selectionMade: false,
  };
}

// Build a deterministic, metadata-only replacement plan for a selected
// presentation image slot. The plan is consumed by the PPJ authoring path;
// this catalog helper never mutates a template or a source-bound PPTX.
export function planTemplateImageReplacement({
  template,
  slotId,
  asset,
  fit = null,
  mask = null,
  accessibility = null,
} = {}) {
  if (template == null || typeof template !== "object" || Array.isArray(template) ||
      template.kind !== "presentation" || typeof template.id !== "string") {
    throw new Error("template must be a presentation search candidate");
  }
  if (typeof slotId !== "string" || !IMAGE_SLOT_ID_PATTERN.test(slotId)) {
    throw new Error("slotId must be a lowercase image-slot identifier");
  }
  const slot = (template.imageSlots ?? []).find((candidate) => candidate.id === slotId);
  if (slot == null) throw new Error(`image slot ${slotId} was not found in ${template.id}`);
  if (asset == null || typeof asset !== "object" || Array.isArray(asset)) {
    throw new Error("asset must be an image asset metadata object");
  }
  assertAssetId(asset.id, "asset.id");
  const widthPx = asset.widthPx;
  const heightPx = asset.heightPx;
  validateReplacementDimension(widthPx, "asset.widthPx");
  validateReplacementDimension(heightPx, "asset.heightPx");
  if (slot.minWidthPx != null && widthPx < slot.minWidthPx) {
    throw new Error(`asset.widthPx must be at least ${slot.minWidthPx} for image slot ${slotId}`);
  }
  if (slot.minHeightPx != null && heightPx < slot.minHeightPx) {
    throw new Error(`asset.heightPx must be at least ${slot.minHeightPx} for image slot ${slotId}`);
  }
  const rights = typeof asset.rights === "string"
    ? asset.rights
    : asset.rights?.status;
  if (slot.rights?.length > 0 && !slot.rights.includes(rights)) {
    throw new Error(`asset.rights must be allowed by image slot ${slotId}`);
  }
  if (asset.sha256 != null) assertHash(asset.sha256, "asset.sha256");
  if (fit != null) {
    assertEnum(fit, "fit", VALID_IMAGE_SLOT_FITS);
    if (slot.allowedFit?.length > 0 && !slot.allowedFit.includes(fit)) {
      throw new Error(`fit ${fit} is not allowed by image slot ${slotId}`);
    }
  }
  if (mask != null) {
    assertEnum(mask, "mask", VALID_IMAGE_SLOT_MASKS);
    if (slot.allowedMask?.length > 0 && !slot.allowedMask.includes(mask)) {
      throw new Error(`mask ${mask} is not allowed by image slot ${slotId}`);
    }
  }
  if (accessibility != null &&
      (typeof accessibility !== "object" || Array.isArray(accessibility))) {
    throw new Error("accessibility must be an object when provided");
  }
  return {
    schema: "office-kit/template-image-slot/v1",
    operation: "replace-image-slot",
    templateId: template.id,
    slotId,
    role: slot.role,
    asset: {
      id: asset.id,
      widthPx,
      heightPx,
      ...(rights == null ? {} : { rights }),
      ...(asset.sha256 == null ? {} : { sha256: asset.sha256 }),
    },
    overrides: {
      ...(fit == null ? {} : { fit }),
      ...(mask == null ? {} : { mask }),
      ...(accessibility == null ? {} : { accessibility: structuredClone(accessibility) }),
    },
    preserve: ["fit", "mask", "crop", "focus", "accessibility"],
    policy: {
      allowedFit: [...(slot.allowedFit ?? [])],
      allowedMask: [...(slot.allowedMask ?? [])],
      minWidthPx: slot.minWidthPx ?? null,
      minHeightPx: slot.minHeightPx ?? null,
      rights: [...(slot.rights ?? [])],
    },
  };
}

// Apply a validated replacement plan to one explicit PPJ image owner.  This
// is deliberately a pure transaction: the caller receives a cloned program
// and may pass it to the PPJ validator/compiler with the matching asset bytes.
// A template slot has no stable identity inside an arbitrary deck, so the
// target element must always be named by the caller instead of being guessed
// from role, z-order, or image content.
export function applyTemplateImageReplacement({
  program,
  plan,
  elementId,
  assetDeclaration = null,
} = {}) {
  assertProgramObject(program);
  validateTemplateImageReplacementPlan(plan);
  assertElementId(elementId, "elementId");

  const root = structuredClone(program);
  if (!Array.isArray(root.pages)) throw new Error("program.pages must be an array");
  const matches = [];
  for (const page of root.pages) collectProgramElements(page?.elements, elementId, matches);
  for (const component of root.components ?? []) {
    collectProgramElements(component?.elements, elementId, matches);
  }
  if (matches.length !== 1) {
    throw new Error(`PPJ image element ${elementId} must resolve exactly once`);
  }
  const target = matches[0];
  if (target?.type !== "image") {
    throw new Error(`PPJ element ${elementId} is not an image`);
  }

  const assets = root.assets;
  if (!Array.isArray(assets)) throw new Error("program.assets must be an array");
  const existingAsset = assets.find((asset) => asset?.id === plan.asset.id) ?? null;
  let addedAsset = false;
  if (existingAsset == null) {
    if (assetDeclaration == null) {
      throw new Error(`assetDeclaration is required for new image asset ${plan.asset.id}`);
    }
    const normalized = normalizeTemplateAssetDeclaration(plan.asset, assetDeclaration);
    assets.push(normalized);
    addedAsset = true;
  } else {
    validateExistingTemplateAsset(plan.asset, existingAsset);
    if (assetDeclaration != null) {
      const normalized = normalizeTemplateAssetDeclaration(plan.asset, assetDeclaration);
      if (!sameAssetIdentity(existingAsset, normalized)) {
        throw new Error(`assetDeclaration for ${plan.asset.id} conflicts with the existing PPJ asset`);
      }
    }
  }

  const sourceBound = target.nativeRef != null && typeof target.nativeRef === "object";
  const changedFields = [];
  if (target.asset !== plan.asset.id) {
    if (sourceBound) requireTemplateCapability(target, "replaceImage", "image.asset");
    target.asset = plan.asset.id;
    changedFields.push("image.asset");
  }

  const overrides = plan.overrides;
  if (Object.hasOwn(overrides, "fit")) {
    const fit = overrides.fit;
    if (target.focus != null && fit !== "cover") {
      throw new Error("image.focus cannot be preserved when a replacement fit is not cover");
    }
    if (sourceBound) requireTemplateCapability(target, "setImageFit", "image.fit");
    if (target.fit !== fit) {
      target.fit = fit;
      changedFields.push("image.fit");
    }
  }

  if (Object.hasOwn(overrides, "mask")) {
    const mask = overrides.mask;
    if (mask === "custom") {
      throw new Error("custom template image masks require explicit PPJ geometry and cannot be synthesized from a slot plan");
    }
    const nextMask = mask === "none" || mask === "rect"
      ? null
      : { kind: "preset", preset: mask };
    if (!sameJsonValue(target.mask ?? null, nextMask)) {
      if (sourceBound) {
        requireTemplateCapability(target, "setImageMask", "image.mask.preset");
        if (target.mask?.kind === "custom") {
          requireTemplateCapability(target, "setImageMask", "image.mask.paths");
        }
        if (target.mask?.kind === "preset" && (target.mask.adjustments?.length ?? 0) > 0) {
          requireTemplateCapability(target, "setImageMask", "image.mask.adjustments");
        }
      }
      if (nextMask == null) delete target.mask;
      else target.mask = nextMask;
      changedFields.push("image.mask.preset");
    }
  }

  if (Object.hasOwn(overrides, "accessibility")) {
    if (sourceBound) {
      throw new Error("source-bound template image accessibility overrides require an explicit accessibility capability");
    }
    if (!sameJsonValue(target.accessibility ?? null, overrides.accessibility)) {
      target.accessibility = structuredClone(overrides.accessibility);
      changedFields.push("accessibility");
    }
  }

  return Object.freeze({
    program: root,
    elementId,
    templateId: plan.templateId,
    slotId: plan.slotId,
    sourceBound,
    addedAsset,
    changedFields: Object.freeze([...new Set(changedFields)]),
  });
}

// Apply a replacement plan and run the resulting PPJ through the native
// presentation compiler and projector.  The native imports stay dynamic so
// template search remains metadata-only and does not initialize OfficeKit at
// module load time.  A caller may inject compile/project functions for a
// harness, but the default path is always the bundled PPJ NativeAOT codec.
export async function applyTemplateImageReplacementToPptx({
  program,
  plan,
  elementId,
  source = new Uint8Array(),
  assetDeclaration = null,
  assetData = null,
  assetDataById = null,
  includeNodeMap = true,
  sourceUri = null,
  assetRootUri = null,
  limits = {},
  compile = null,
  project = null,
} = {}) {
  const applied = applyTemplateImageReplacement({
    program,
    plan,
    elementId,
    assetDeclaration,
  });
  const sourceBytes = copyBinary(source, "PPTX source");
  const sourceDescriptor = applied.program.source;
  if (sourceDescriptor != null) {
    if (sourceBytes.byteLength === 0) {
      throw new Error("source-bound template replacement requires the exact PPTX source bytes");
    }
    if (typeof sourceDescriptor.sha256 === "string" &&
        hashBytes(sourceBytes) !== sourceDescriptor.sha256) {
      throw new Error("source-bound template replacement source bytes do not match program.source.sha256");
    }
  } else if (sourceBytes.byteLength !== 0) {
    throw new Error("source-free template replacement cannot attach a PPTX source package");
  }

  const replacementAsset = applied.program.assets.find((asset) => asset?.id === plan.asset.id);
  if (replacementAsset == null) {
    throw new Error(`PPJ replacement asset ${plan.asset.id} is missing after applying the plan`);
  }
  const nativeAssets = [];
  for (const declaration of applied.program.assets) {
    const rawData = declaration.id === replacementAsset.id
      ? assetData
      : lookupAssetData(assetDataById, declaration.id);
    if (rawData == null) {
      if (declaration.id === replacementAsset.id && applied.addedAsset) {
        throw new Error(`assetData is required for new image asset ${replacementAsset.id}`);
      }
      // Source-bound compilation rehydrates baseline assets from the exact
      // PPTX package. Source-free compilation will reject any omitted
      // declaration in the native compiler; keeping the omission here lets a
      // caller use a fail-closed injected harness without fabricating bytes.
      continue;
    }
    const data = copyBinary(rawData, `PPJ asset ${declaration.id}`);
    if (data.byteLength === 0) throw new Error(`PPJ asset ${declaration.id} must contain bytes`);
    const digest = hashBytes(data);
    if (digest !== declaration.sha256) {
      throw new Error(`PPJ asset ${declaration.id} bytes do not match its declared SHA-256`);
    }
    nativeAssets.push({
      id: declaration.id,
      fileName: declaration.uri,
      mimeType: declaration.mimeType,
      sha256: declaration.sha256,
      data,
    });
  }

  const nativeCompile = compile ?? (await import("../ppj/native.mjs")).compilePpjToPptx;
  if (typeof nativeCompile !== "function") throw new TypeError("compile must be a function");
  const compiled = await nativeCompile(
    Buffer.from(JSON.stringify(applied.program), "utf8"),
    {
      source: sourceBytes,
      assets: nativeAssets,
      includeNodeMap: Boolean(includeNodeMap),
      limits,
    },
  );
  const compileReceipt = validateTemplateCompileReceipt(compiled);

  const nativeProject = project ?? (await import("../ppj/native.mjs")).projectPptxToPpj;
  if (typeof nativeProject !== "function") throw new TypeError("project must be a function");
  const projectionSourceUri = sourceUri ?? sourceDescriptor?.uri ??
    "deck.assets/source/template-image-replacement.pptx";
  const projectionAssetRootUri = assetRootUri ?? inferAssetRootUri(applied.program.assets);
  assertRelativeAssetPath(projectionSourceUri, "sourceUri");
  assertRelativeAssetPath(projectionAssetRootUri, "assetRootUri");
  const reprojected = validateTemplateProjectionReceipt(await nativeProject(
    compileReceipt.file,
    {
      sourceUri: projectionSourceUri,
      assetRootUri: projectionAssetRootUri,
      includeNodeMap: Boolean(includeNodeMap),
      limits,
    },
  ));
  if (sourceDescriptor != null && reprojected.sourceBound !== true) {
    throw new Error("source-bound template replacement did not remain source-bound after reprojection");
  }

  return Object.freeze({
    ...applied,
    compile: compileReceipt,
    reproject: reprojected,
  });
}

function copyBinary(value, label) {
  if (value instanceof Uint8Array) return Uint8Array.from(value);
  if (Buffer.isBuffer(value)) return Uint8Array.from(value);
  throw new TypeError(`${label} must be a Uint8Array.`);
}

function lookupAssetData(assetDataById, id) {
  if (assetDataById == null) return null;
  if (assetDataById instanceof Map) return assetDataById.get(id) ?? null;
  if (typeof assetDataById !== "object" || Array.isArray(assetDataById)) {
    throw new TypeError("assetDataById must be a Map or an object keyed by PPJ asset id");
  }
  return Object.hasOwn(assetDataById, id) ? assetDataById[id] : null;
}

function hashBytes(value) {
  return createHash("sha256").update(value).digest("hex");
}

function validateTemplateCompileReceipt(value) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("native template replacement compiler returned an invalid receipt");
  }
  const file = copyBinary(value.file, "native template replacement PPTX");
  if (file.byteLength === 0) throw new Error("native template replacement compiler returned an empty PPTX");
  if (typeof value.outputSha256 !== "string" || !HASH_PATTERN.test(value.outputSha256) ||
      hashBytes(file) !== value.outputSha256) {
    throw new Error("native template replacement compiler returned an invalid output hash");
  }
  if (!Array.isArray(value.changedParts) ||
      !value.changedParts.every((part) => typeof part === "string" && part.length > 0)) {
    throw new Error("native template replacement compiler returned invalid changed parts");
  }
  return Object.freeze({ ...value, file });
}

function validateTemplateProjectionReceipt(value) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("native template replacement projector returned an invalid receipt");
  }
  const programJson = copyBinary(value.programJson, "reprojected PPJ");
  if (programJson.byteLength === 0) throw new Error("native template replacement projector returned an empty PPJ");
  return Object.freeze({ ...value, programJson });
}

function inferAssetRootUri(assets) {
  for (const asset of assets ?? []) {
    if (typeof asset?.uri !== "string") continue;
    const separator = asset.uri.lastIndexOf("/");
    if (separator > 0) return asset.uri.slice(0, separator);
  }
  return "deck.assets/media";
}

function assertProgramObject(value) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("program must be a PPJ object");
  }
}

function assertElementId(value, label) {
  if (typeof value !== "string" || value.trim() !== value || value.length === 0 || value.length > 512 || /[\0\r\n]/u.test(value)) {
    throw new Error(`${label} must be a bounded non-empty element identifier`);
  }
}

function validateTemplateImageReplacementPlan(plan) {
  if (plan == null || typeof plan !== "object" || Array.isArray(plan)) {
    throw new Error("plan must be a template image replacement plan");
  }
  assertObjectKeys(plan, "plan", [
    "schema",
    "operation",
    "templateId",
    "slotId",
    "role",
    "asset",
    "overrides",
    "preserve",
    "policy",
  ]);
  if (plan.schema !== "office-kit/template-image-slot/v1" || plan.operation !== "replace-image-slot") {
    throw new Error("plan must use office-kit/template-image-slot/v1 replace-image-slot semantics");
  }
  assertTemplateId(plan.templateId, "plan.templateId");
  if (typeof plan.slotId !== "string" || !IMAGE_SLOT_ID_PATTERN.test(plan.slotId)) {
    throw new Error("plan.slotId must be a lowercase image-slot identifier");
  }
  assertEnum(plan.role, "plan.role", VALID_IMAGE_SLOT_ROLES);
  if (plan.asset == null || typeof plan.asset !== "object" || Array.isArray(plan.asset)) {
    throw new Error("plan.asset must be an image asset metadata object");
  }
  assertObjectKeys(plan.asset, "plan.asset", ["id", "widthPx", "heightPx", "rights", "sha256"]);
  assertAssetId(plan.asset.id, "plan.asset.id");
  validateReplacementDimension(plan.asset.widthPx, "plan.asset.widthPx");
  validateReplacementDimension(plan.asset.heightPx, "plan.asset.heightPx");
  if (plan.asset.rights != null) assertEnum(plan.asset.rights, "plan.asset.rights", VALID_IMAGE_SLOT_RIGHTS);
  if (plan.asset.sha256 != null) assertHash(plan.asset.sha256, "plan.asset.sha256");
  if (plan.overrides == null || typeof plan.overrides !== "object" || Array.isArray(plan.overrides)) {
    throw new Error("plan.overrides must be an object");
  }
  assertObjectKeys(plan.overrides, "plan.overrides", ["fit", "mask", "accessibility"]);
  if (plan.overrides.fit != null) assertEnum(plan.overrides.fit, "plan.overrides.fit", VALID_IMAGE_SLOT_FITS);
  if (plan.overrides.mask != null) assertEnum(plan.overrides.mask, "plan.overrides.mask", VALID_IMAGE_SLOT_MASKS);
  if (plan.overrides.accessibility != null) validateTemplateAccessibility(plan.overrides.accessibility);
  if (!Array.isArray(plan.preserve) || plan.preserve.length > 16 ||
      !plan.preserve.every((field) => typeof field === "string" && field.length > 0)) {
    throw new Error("plan.preserve must be a bounded field list");
  }
  if (plan.policy == null || typeof plan.policy !== "object" || Array.isArray(plan.policy)) {
    throw new Error("plan.policy must be an object");
  }
  assertObjectKeys(plan.policy, "plan.policy", ["allowedFit", "allowedMask", "minWidthPx", "minHeightPx", "rights"]);
  for (const field of ["allowedFit", "allowedMask", "rights"]) {
    if (!Array.isArray(plan.policy[field])) throw new Error(`plan.policy.${field} must be an array`);
  }
  validateBoundedEnumArray(plan.policy.allowedFit, "plan.policy.allowedFit", VALID_IMAGE_SLOT_FITS, 3, { allowEmpty: true });
  validateBoundedEnumArray(plan.policy.allowedMask, "plan.policy.allowedMask", VALID_IMAGE_SLOT_MASKS, 16, { allowEmpty: true });
  validateBoundedEnumArray(plan.policy.rights, "plan.policy.rights", VALID_IMAGE_SLOT_RIGHTS, 16, { allowEmpty: true });
  validateOptionalPixelDimension(plan.policy.minWidthPx, "plan.policy.minWidthPx");
  validateOptionalPixelDimension(plan.policy.minHeightPx, "plan.policy.minHeightPx");
  if (plan.policy.minWidthPx != null && plan.asset.widthPx < plan.policy.minWidthPx) {
    throw new Error("plan.asset.widthPx does not satisfy plan.policy.minWidthPx");
  }
  if (plan.policy.minHeightPx != null && plan.asset.heightPx < plan.policy.minHeightPx) {
    throw new Error("plan.asset.heightPx does not satisfy plan.policy.minHeightPx");
  }
  if (plan.overrides.fit != null && plan.policy.allowedFit.length > 0 && !plan.policy.allowedFit.includes(plan.overrides.fit)) {
    throw new Error("plan.overrides.fit is not allowed by plan.policy.allowedFit");
  }
  if (plan.overrides.mask != null && plan.policy.allowedMask.length > 0 && !plan.policy.allowedMask.includes(plan.overrides.mask)) {
    throw new Error("plan.overrides.mask is not allowed by plan.policy.allowedMask");
  }
  if (plan.asset.rights != null && plan.policy.rights.length > 0 && !plan.policy.rights.includes(plan.asset.rights)) {
    throw new Error("plan.asset.rights is not allowed by plan.policy.rights");
  }
}

function collectProgramElements(elements, id, output) {
  for (const element of elements ?? []) {
    if (element?.id === id) output.push(element);
    collectProgramElements(element?.elements, id, output);
    collectProgramElements(element?.children, id, output);
  }
}

function normalizeTemplateAssetDeclaration(planAsset, declaration) {
  if (declaration == null || typeof declaration !== "object" || Array.isArray(declaration)) {
    throw new Error("assetDeclaration must be an object");
  }
  assertObjectKeys(declaration, "assetDeclaration", [
    "id",
    "uri",
    "mimeType",
    "sha256",
    "widthPx",
    "heightPx",
    "rights",
    "accessibility",
  ]);
  if (declaration.id !== planAsset.id) throw new Error("assetDeclaration.id must match plan.asset.id");
  assertRelativeAssetPath(declaration.uri, "assetDeclaration.uri");
  if (typeof declaration.mimeType !== "string" || !/^[A-Za-z0-9!#$&^_.+-]+\/[A-Za-z0-9!#$&^_.+-]+$/u.test(declaration.mimeType)) {
    throw new Error("assetDeclaration.mimeType must be a valid MIME type");
  }
  assertHash(declaration.sha256, "assetDeclaration.sha256");
  if (planAsset.sha256 != null && declaration.sha256 !== planAsset.sha256) {
    throw new Error("assetDeclaration.sha256 must match plan.asset.sha256");
  }
  const widthPx = declaration.widthPx ?? planAsset.widthPx;
  const heightPx = declaration.heightPx ?? planAsset.heightPx;
  validateReplacementDimension(widthPx, "assetDeclaration.widthPx");
  validateReplacementDimension(heightPx, "assetDeclaration.heightPx");
  if (widthPx !== planAsset.widthPx || heightPx !== planAsset.heightPx) {
    throw new Error("assetDeclaration dimensions must match plan.asset metadata");
  }
  if (declaration.rights == null || typeof declaration.rights !== "object" || Array.isArray(declaration.rights)) {
    throw new Error("assetDeclaration.rights must be an object");
  }
  if (typeof declaration.rights.status !== "string") throw new Error("assetDeclaration.rights.status is required");
  if (planAsset.rights != null && declaration.rights.status !== planAsset.rights) {
    throw new Error("assetDeclaration.rights.status must match plan.asset.rights");
  }
  validateTemplateAccessibility(declaration.accessibility);
  return {
    ...structuredClone(declaration),
    widthPx,
    heightPx,
  };
}

function validateExistingTemplateAsset(planAsset, existing) {
  if (existing == null || typeof existing !== "object" || Array.isArray(existing)) {
    throw new Error(`PPJ asset ${planAsset.id} is invalid`);
  }
  if (existing.sha256 != null && planAsset.sha256 != null && existing.sha256 !== planAsset.sha256) {
    throw new Error(`PPJ asset ${planAsset.id} does not match plan.asset.sha256`);
  }
  if (existing.widthPx != null && existing.widthPx !== planAsset.widthPx) {
    throw new Error(`PPJ asset ${planAsset.id} does not match plan.asset.widthPx`);
  }
  if (existing.heightPx != null && existing.heightPx !== planAsset.heightPx) {
    throw new Error(`PPJ asset ${planAsset.id} does not match plan.asset.heightPx`);
  }
  const rights = existing.rights?.status;
  if (planAsset.rights != null && rights != null && rights !== planAsset.rights) {
    throw new Error(`PPJ asset ${planAsset.id} does not match plan.asset.rights`);
  }
}

function sameAssetIdentity(left, right) {
  return left.id === right.id && left.uri === right.uri && left.mimeType === right.mimeType &&
    left.sha256 === right.sha256 && left.widthPx === right.widthPx && left.heightPx === right.heightPx &&
    sameJsonValue(left.rights, right.rights) && sameJsonValue(left.accessibility, right.accessibility);
}

function validateTemplateAccessibility(value) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("accessibility must be an object");
  }
  assertObjectKeys(value, "accessibility", ["decorative", "title", "description"]);
  if (typeof value.decorative !== "boolean") throw new Error("accessibility.decorative must be a boolean");
  if (value.title != null) assertShortString(value.title, "accessibility.title", 512);
  if (value.description != null) assertShortString(value.description, "accessibility.description", 2048);
  if (value.decorative && (value.title != null || value.description != null)) {
    throw new Error("decorative accessibility cannot include title or description");
  }
}

function requireTemplateCapability(element, operation, field) {
  const capabilities = element.nativeRef?.capabilities;
  if (!Array.isArray(capabilities)) {
    throw new Error(`source-bound image ${element.id} does not expose ${operation}.${field}`);
  }
  const supported = capabilities.some((capability) =>
    capability?.operation === operation && Array.isArray(capability.fields) && capability.fields.includes(field));
  if (!supported) throw new Error(`source-bound image ${element.id} does not expose ${operation}.${field}`);
}

function sameJsonValue(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

async function resolveRoots(explicitRoots, projectPath) {
  if (explicitRoots != null && !Array.isArray(explicitRoots)) {
    throw new Error("Template roots must be an array of paths.");
  }
  if (explicitRoots != null && explicitRoots.length > 20) {
    throw new Error("At most 20 template roots may be queried.");
  }
  const requested = explicitRoots == null || explicitRoots.length === 0
    ? await defaultRoots(projectPath)
    : explicitRoots.map((root) => ({ path: root, source: "explicit" }));
  const resolved = [];
  const seen = new Set();

  for (const entry of requested) {
    if (typeof entry.path !== "string" || entry.path.trim().length === 0) {
      throw new Error("Template roots must be non-empty paths.");
    }
    const absolutePath = path.resolve(entry.path);
    const stat = await fs.lstat(absolutePath).catch((error) => {
      if (error?.code === "ENOENT" && entry.source !== "explicit") return null;
      throw error;
    });
    if (stat == null) continue;
    if (!stat.isDirectory() || stat.isSymbolicLink()) {
      throw new Error(`Template root must be a real directory: ${absolutePath}`);
    }
    const canonicalPath = await fs.realpath(absolutePath);
    if (seen.has(canonicalPath)) continue;
    seen.add(canonicalPath);
    resolved.push({ path: canonicalPath, source: entry.source });
  }
  return resolved;
}

async function defaultRoots(projectPath) {
  const configured = (process.env.OFFICE_KIT_TEMPLATE_ROOTS ?? "")
    .split(path.delimiter)
    .filter(Boolean)
    .map((root) => ({ path: root, source: "configured" }));
  const officeKitHome = process.env.OFFICE_KIT_HOME == null
    ? path.join(os.homedir(), ".office-kit")
    : path.resolve(process.env.OFFICE_KIT_HOME);
  return [
    ...configured,
    ...await projectTemplateRoots(projectPath),
    { path: path.join(officeKitHome, "skills"), source: "local-user" },
    {
      path: path.join(PACKAGE_ROOT, "skills/presentation-template-library/skills"),
      source: "package-default",
    },
    {
      path: path.join(PACKAGE_ROOT, "skills/default-template-library/skills"),
      source: "package-default",
    },
  ];
}

async function projectTemplateRoots(projectPath) {
  const absoluteProject = path.resolve(projectPath);
  const manifestPath = path.join(absoluteProject, MANAGED_SKILLS_MANIFEST);
  const stat = await fs.lstat(manifestPath).catch((error) => {
    if (error?.code === "ENOENT") return null;
    throw error;
  });
  if (stat == null) return [];
  if (!stat.isFile() || stat.isSymbolicLink() || stat.size > MAX_SIDECAR_BYTES) {
    throw new Error(`${MANAGED_SKILLS_MANIFEST} must be a bounded regular file.`);
  }
  let manifest;
  try {
    manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  } catch (error) {
    throw new Error(`${MANAGED_SKILLS_MANIFEST} is not valid JSON: ${error.message}`);
  }
  if (manifest?.schemaVersion !== 1 || !Array.isArray(manifest.installations)) {
    throw new Error(`${MANAGED_SKILLS_MANIFEST} uses an unsupported or invalid schema.`);
  }
  const roots = [];
  const seen = new Set();
  for (const installation of manifest.installations) {
    const relativePath = installation?.path;
    if (
      typeof relativePath !== "string" ||
      path.isAbsolute(relativePath) ||
      relativePath.includes("\\") ||
      relativePath.split("/").some((segment) => segment === "" || segment === "." || segment === "..")
    ) {
      throw new Error(`${MANAGED_SKILLS_MANIFEST} contains an unsafe installation path.`);
    }
    const skillsRoot = path.resolve(absoluteProject, path.posix.dirname(relativePath));
    if (!isInsideOrEqual(absoluteProject, skillsRoot) || seen.has(skillsRoot)) continue;
    seen.add(skillsRoot);
    roots.push({ path: skillsRoot, source: "project" });
  }
  return roots;
}

async function readTemplate({ expectedId, root, templatePath }) {
  const templateStat = await fs.lstat(templatePath);
  if (!templateStat.isDirectory() || templateStat.isSymbolicLink()) {
    throw new Error("template root must be a real directory");
  }
  const sidecarPath = path.join(templatePath, SIDECAR_NAME);
  const sidecarStat = await fs.lstat(sidecarPath);
  if (!sidecarStat.isFile() || sidecarStat.isSymbolicLink()) {
    throw new Error(`${SIDECAR_NAME} must be a regular file`);
  }
  if (sidecarStat.size > MAX_SIDECAR_BYTES) {
    throw new Error(`${SIDECAR_NAME} exceeds the ${MAX_SIDECAR_BYTES}-byte budget`);
  }

  let metadata;
  try {
    metadata = JSON.parse(await fs.readFile(sidecarPath, "utf8"));
  } catch (error) {
    throw new Error(`${SIDECAR_NAME} is not valid JSON: ${error.message}`);
  }
  validateMetadata(metadata, expectedId);

  const skillPath = await resolveTemplateSkill(
    templatePath,
    metadata.kind === "presentation" ? metadata.provenance.guideSha256 : null,
  );
  const previewPath = await resolveAsset(
    templatePath,
    metadata.preview,
    metadata.provenance.previewSha256,
    "preview",
  );

  const shared = {
    templateSchemaVersion: metadata.schemaVersion,
    id: metadata.id,
    displayName: metadata.displayName,
    kind: metadata.kind,
    useWhen: metadata.useWhen,
    avoidWhen: metadata.avoidWhen,
    audiences: metadata.audiences,
    contentShapes: metadata.contentShapes,
    visualTraits: metadata.visualTraits,
    visualCommitment: metadata.visualCommitment,
    provenance: { ...metadata.provenance },
    catalogSource: root.source,
    templateRoot: await fs.realpath(templatePath),
    skillPath,
    previewPath,
  };
  if (metadata.kind === "presentation") {
    await assertPresentationTemplateSurface(templatePath, metadata);
    const examplePaths = await Promise.all(
      metadata.examples.map((example, index) =>
        resolveAsset(
          templatePath,
          example.path,
          example.sha256,
          `examples[${index}]`,
        )),
    );
    const referenceProgramPath = metadata.referenceProgram == null ? null : await resolveOptionalAsset(
      templatePath,
      metadata.referenceProgram,
      "referenceProgram",
    );
    const referencePptxPath = metadata.referencePptx == null ? null : await resolveOptionalAsset(
      templatePath,
      metadata.referencePptx,
      "referencePptx",
    );
    const imageSlots = (metadata.imageSlots ?? []).map((slot) => {
      const exampleIndex = metadata.examples.findIndex((example) => example.path === slot.examplePath);
      if (exampleIndex < 0) throw new Error(`imageSlots.${slot.id} example binding is missing`);
      const example = metadata.examples[exampleIndex];
      return {
        ...slot,
        allowedFit: slot.allowedFit == null ? [] : [...slot.allowedFit],
        allowedMask: slot.allowedMask == null ? [] : [...slot.allowedMask],
        rights: slot.rights == null ? [] : [...slot.rights],
        example: {
          path: example.path,
          sha256: example.sha256,
          absolutePath: examplePaths[exampleIndex],
        },
      };
    });
    return {
      ...shared,
      imageSlots,
      examples: metadata.examples.map((example, index) => ({
        role: example.role,
        path: example.path,
        sha256: example.sha256,
        absolutePath: examplePaths[index],
      })),
      examplePaths,
      referenceProgram: metadata.referenceProgram == null ? null : {
        ...metadata.referenceProgram,
        absolutePath: referenceProgramPath,
        available: referenceProgramPath != null,
        fetchCommand: referenceProgramPath == null ? `officekit template fetch ${metadata.id}` : null,
      },
      referencePptx: metadata.referencePptx == null ? null : {
        ...metadata.referencePptx,
        absolutePath: referencePptxPath,
        available: referencePptxPath != null,
        fetchCommand: referencePptxPath == null ? `officekit template fetch ${metadata.id}` : null,
      },
    };
  }

  const referencePath = await resolveAsset(
    templatePath,
    metadata.reference,
    metadata.provenance.referenceSha256,
    "reference",
  );
  return {
    ...shared,
    editProfile: metadata.editProfile,
    referencePath,
  };
}

function validateMetadata(value, expectedId) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("metadata must be an object");
  }
  if (value.schemaVersion === 2 && value.kind === "presentation") {
    throw new Error(
      "presentation schema v2 is unsupported; rebuild it with presentation-template-creator",
    );
  }
  if (value.schemaVersion === 3) {
    validatePresentationMetadata(value, expectedId);
    return;
  }
  if (value.schemaVersion !== 2) {
    throw new Error("schemaVersion must be 2 for document/spreadsheet or 3 for presentation");
  }
  validateSourceTemplateMetadata(value, expectedId);
}

function validateSourceTemplateMetadata(value, expectedId) {
  assertObjectKeys(
    value,
    "metadata",
    [
      "schemaVersion",
      "id",
      "displayName",
      "kind",
      "reference",
      "preview",
      "useWhen",
      "avoidWhen",
      "audiences",
      "contentShapes",
      "visualTraits",
      "visualCommitment",
      "editProfile",
      "provenance",
    ],
  );
  assertTemplateId(value.id, "id");
  if (value.id !== expectedId) {
    throw new Error(`id must match directory name ${expectedId}`);
  }
  assertShortString(value.displayName, "displayName", 80);
  assertKind(value.kind);
  if (value.kind === "presentation") {
    throw new Error(
      "presentation schema v2 is unsupported; rebuild it with presentation-template-creator",
    );
  }
  assertEnglishSearchArray(value.useWhen, "useWhen", { min: 1, max: 20 });
  assertEnglishSearchArray(value.avoidWhen, "avoidWhen", { min: 0, max: 20 });
  assertEnglishSearchArray(value.audiences, "audiences", { min: 0, max: 20 });
  assertEnglishSearchArray(value.contentShapes, "contentShapes", { min: 0, max: 20 });
  assertRelativeAssetPath(value.reference, "reference");
  assertRelativeAssetPath(value.preview, "preview");
  if (path.posix.extname(value.reference).toLowerCase() !== REFERENCE_EXTENSIONS.get(value.kind)) {
    throw new Error(`${value.kind} templates must use a ${REFERENCE_EXTENSIONS.get(value.kind)} reference`);
  }
  if (path.posix.extname(value.preview).toLowerCase() !== ".png") {
    throw new Error("preview must use a .png file");
  }

  if (value.visualTraits == null || typeof value.visualTraits !== "object" || Array.isArray(value.visualTraits)) {
    throw new Error("visualTraits must be an object");
  }
  assertObjectKeys(
    value.visualTraits,
    "visualTraits",
    ["tone", "density", "colorMode", "structure"],
  );
  assertEnglishSearchArray(value.visualTraits.tone, "visualTraits.tone", { min: 0, max: 12 });
  assertEnum(value.visualTraits.density, "visualTraits.density", VALID_DENSITIES);
  assertEnum(value.visualTraits.colorMode, "visualTraits.colorMode", VALID_COLOR_MODES);
  assertEnglishSearchArray(value.visualTraits.structure, "visualTraits.structure", { min: 0, max: 12 });
  assertEnum(value.visualCommitment, "visualCommitment", VALID_COMMITMENTS);

  if (value.editProfile == null || typeof value.editProfile !== "object" || Array.isArray(value.editProfile)) {
    throw new Error("editProfile must be an object");
  }
  assertObjectKeys(
    value.editProfile,
    "editProfile",
    ["level", "verifiedOperations"],
  );
  assertEnum(value.editProfile.level, "editProfile.level", VALID_EDIT_LEVELS);
  assertStringArray(value.editProfile.verifiedOperations, "editProfile.verifiedOperations", {
    min: 0,
    max: 24,
  });
  if (value.editProfile.level === "copy-only" && value.editProfile.verifiedOperations.length !== 0) {
    throw new Error("copy-only templates cannot declare verifiedOperations");
  }

  if (value.provenance == null || typeof value.provenance !== "object" || Array.isArray(value.provenance)) {
    throw new Error("provenance must be an object");
  }
  assertObjectKeys(
    value.provenance,
    "provenance",
    ["license", "source", "referenceSha256", "previewSha256"],
  );
  assertShortString(value.provenance.license, "provenance.license", 120);
  assertShortString(value.provenance.source, "provenance.source", 500);
  assertHash(value.provenance.referenceSha256, "provenance.referenceSha256");
  assertHash(value.provenance.previewSha256, "provenance.previewSha256");
}

function validatePresentationMetadata(value, expectedId) {
  assertObjectKeys(
    value,
    "metadata",
    [
      "schemaVersion",
      "id",
      "displayName",
      "kind",
      "preview",
      "examples",
      "useWhen",
      "avoidWhen",
      "audiences",
      "contentShapes",
      "visualTraits",
      "visualCommitment",
      "imageSlots",
      "referenceProgram",
      "referencePptx",
      "provenance",
    ],
  );
  if (value.kind !== "presentation") {
    throw new Error("schemaVersion 3 is reserved for presentation templates");
  }
  assertTemplateId(value.id, "id");
  if (value.id !== expectedId) {
    throw new Error(`id must match directory name ${expectedId}`);
  }
  assertShortString(value.displayName, "displayName", 80);
  assertEnglishSearchArray(value.useWhen, "useWhen", { min: 1, max: 20 });
  assertEnglishSearchArray(value.avoidWhen, "avoidWhen", { min: 0, max: 20 });
  assertEnglishSearchArray(value.audiences, "audiences", { min: 0, max: 20 });
  assertEnglishSearchArray(value.contentShapes, "contentShapes", { min: 0, max: 20 });
  validateVisualMetadata(value);
  assertRelativePngPath(value.preview, "preview");
  if (!Array.isArray(value.examples) ||
      value.examples.length < MIN_PRESENTATION_EXAMPLES ||
      value.examples.length > MAX_PRESENTATION_EXAMPLES) {
    throw new Error(
      `examples must contain ${MIN_PRESENTATION_EXAMPLES}-${MAX_PRESENTATION_EXAMPLES} entries`,
    );
  }
  const paths = new Set();
  const roles = new Set();
  for (const [index, example] of value.examples.entries()) {
    if (example == null || typeof example !== "object" || Array.isArray(example)) {
      throw new Error(`examples[${index}] must be an object`);
    }
    assertObjectKeys(example, `examples[${index}]`, ["path", "role", "sha256"]);
    assertRelativePngPath(example.path, `examples[${index}].path`);
    if (!example.path.startsWith("assets/examples/")) {
      throw new Error(`examples[${index}].path must be under assets/examples/`);
    }
    assertEnum(example.role, `examples[${index}].role`, VALID_PRESENTATION_EXAMPLE_ROLES);
    assertHash(example.sha256, `examples[${index}].sha256`);
    if (paths.has(example.path)) throw new Error("examples must use unique paths");
    paths.add(example.path);
    roles.add(example.role);
  }
  if (roles.size < 3) throw new Error("examples must cover at least 3 distinct roles");
  validatePresentationImageSlots(value.imageSlots, value.examples);
  validatePresentationReference(value.referenceProgram, "referenceProgram", ".ppj");
  validatePresentationReference(value.referencePptx, "referencePptx", ".pptx");
  if (value.provenance == null || typeof value.provenance !== "object" || Array.isArray(value.provenance)) {
    throw new Error("provenance must be an object");
  }
  assertObjectKeys(
    value.provenance,
    "provenance",
    ["license", "source", "guideSha256", "previewSha256"],
  );
  assertShortString(value.provenance.license, "provenance.license", 120);
  assertShortString(value.provenance.source, "provenance.source", 500);
  assertHash(value.provenance.guideSha256, "provenance.guideSha256");
  assertHash(value.provenance.previewSha256, "provenance.previewSha256");
}

function validatePresentationImageSlots(value, examples) {
  if (value == null) return;
  if (!Array.isArray(value) || value.length > 64) {
    throw new Error("imageSlots must contain 0-64 entries");
  }
  const exampleByPath = new Map(examples.map((example) => [example.path, example]));
  const ids = new Set();
  for (const [index, slot] of value.entries()) {
    if (slot == null || typeof slot !== "object" || Array.isArray(slot)) {
      throw new Error(`imageSlots[${index}] must be an object`);
    }
    assertObjectKeys(slot, `imageSlots[${index}]`, [
      "id",
      "role",
      "examplePath",
      "allowedFit",
      "allowedMask",
      "minWidthPx",
      "minHeightPx",
      "rights",
    ]);
    if (typeof slot.id !== "string" || !IMAGE_SLOT_ID_PATTERN.test(slot.id)) {
      throw new Error(`imageSlots[${index}].id must be a lowercase identifier`);
    }
    if (ids.has(slot.id)) throw new Error(`imageSlots must use unique ids: ${slot.id}`);
    ids.add(slot.id);
    assertEnum(slot.role, `imageSlots[${index}].role`, VALID_IMAGE_SLOT_ROLES);
    if (typeof slot.examplePath !== "string" || !exampleByPath.has(slot.examplePath)) {
      throw new Error(`imageSlots[${index}].examplePath must reference a declared example`);
    }
    validateBoundedEnumArray(slot.allowedFit, `imageSlots[${index}].allowedFit`, VALID_IMAGE_SLOT_FITS, 3);
    validateBoundedEnumArray(slot.allowedMask, `imageSlots[${index}].allowedMask`, VALID_IMAGE_SLOT_MASKS, 16);
    validateBoundedEnumArray(slot.rights, `imageSlots[${index}].rights`, VALID_IMAGE_SLOT_RIGHTS, 16);
    validateOptionalPixelDimension(slot.minWidthPx, `imageSlots[${index}].minWidthPx`);
    validateOptionalPixelDimension(slot.minHeightPx, `imageSlots[${index}].minHeightPx`);
  }
}

function validateBoundedEnumArray(value, label, allowed, max, { allowEmpty = false } = {}) {
  if (value == null) return;
  const minimum = allowEmpty ? 0 : 1;
  if (!Array.isArray(value) || value.length < minimum || value.length > max) {
    throw new Error(`${label} must contain ${minimum}-${max} values`);
  }
  const seen = new Set();
  for (const item of value) {
    assertEnum(item, label, allowed);
    if (seen.has(item)) throw new Error(`${label} must not contain duplicates`);
    seen.add(item);
  }
}

function validateOptionalPixelDimension(value, label) {
  if (value == null) return;
  if (!Number.isSafeInteger(value) || value < 1 || value > 16_384) {
    throw new Error(`${label} must be an integer from 1 to 16384`);
  }
}

function validatePresentationReference(value, label, extension) {
  if (value == null) return;
  if (typeof value !== "object" || Array.isArray(value)) throw new Error(`${label} must be an object`);
  assertObjectKeys(value, label, ["path", "sha256", "license", "source", "download"]);
  assertRelativeAssetPath(value.path, `${label}.path`);
  if (!value.path.startsWith("assets/references/")) throw new Error(`${label}.path must be under assets/references/`);
  if (path.posix.extname(value.path).toLowerCase() !== extension) throw new Error(`${label}.path must use ${extension}`);
  assertHash(value.sha256, `${label}.sha256`);
  assertShortString(value.license, `${label}.license`, 120);
  assertShortString(value.source, `${label}.source`, 500);
  if (value.download != null) validateRemoteReference(value.download, `${label}.download`, value.sha256);
}

function validateRemoteReference(value, label, expectedHash) {
  if (typeof value !== "object" || Array.isArray(value)) throw new Error(`${label} must be an object`);
  assertObjectKeys(value, label, ["url", "sha256", "bytes"]);
  if (typeof value.url !== "string" || value.url.length > 2048 || /[\0\r\n]/u.test(value.url)) {
    throw new Error(`${label}.url must be a bounded HTTPS URL`);
  }
  let parsed;
  try { parsed = new URL(value.url); } catch { throw new Error(`${label}.url must be a valid HTTPS URL`); }
  if (parsed.protocol !== "https:" || parsed.username || parsed.password || parsed.search || parsed.hash ||
      !REMOTE_REFERENCE_HOSTS.has(parsed.hostname.toLowerCase()) || parsed.pathname.includes("..")) {
    throw new Error(`${label}.url must be an HTTPS raw.githubusercontent.com URL without credentials or traversal`);
  }
  assertHash(value.sha256, `${label}.sha256`);
  if (value.sha256 !== expectedHash) throw new Error(`${label}.sha256 must match the declared reference hash`);
  if (!Number.isSafeInteger(value.bytes) || value.bytes < 1 || value.bytes > MAX_REMOTE_REFERENCE_BYTES) {
    throw new Error(`${label}.bytes must be an integer from 1 to ${MAX_REMOTE_REFERENCE_BYTES}`);
  }
}

function validateVisualMetadata(value) {
  if (value.visualTraits == null || typeof value.visualTraits !== "object" || Array.isArray(value.visualTraits)) {
    throw new Error("visualTraits must be an object");
  }
  assertObjectKeys(
    value.visualTraits,
    "visualTraits",
    ["tone", "density", "colorMode", "structure"],
  );
  assertEnglishSearchArray(value.visualTraits.tone, "visualTraits.tone", { min: 0, max: 12 });
  assertEnum(value.visualTraits.density, "visualTraits.density", VALID_DENSITIES);
  assertEnum(value.visualTraits.colorMode, "visualTraits.colorMode", VALID_COLOR_MODES);
  assertEnglishSearchArray(value.visualTraits.structure, "visualTraits.structure", { min: 0, max: 12 });
  assertEnum(value.visualCommitment, "visualCommitment", VALID_COMMITMENTS);
}

async function resolveAsset(templatePath, relativePath, expectedHash, label) {
  const resolved = path.resolve(templatePath, relativePath);
  const canonicalTemplatePath = await fs.realpath(templatePath);
  const stat = await fs.lstat(resolved);
  if (!stat.isFile() || stat.isSymbolicLink()) {
    throw new Error(`${label} must be a regular nonsymlink file`);
  }
  const canonicalAssetPath = await fs.realpath(resolved);
  if (!isInside(canonicalTemplatePath, canonicalAssetPath)) {
    throw new Error(`${label} escapes the template directory`);
  }
  const actualHash = await sha256File(canonicalAssetPath);
  if (actualHash !== expectedHash) {
    throw new Error(`${label} SHA-256 mismatch`);
  }
  return canonicalAssetPath;
}

async function resolveOptionalAsset(templatePath, declaration, label) {
  try {
    return await resolveAsset(templatePath, declaration.path, declaration.sha256, label);
  } catch (error) {
    if (declaration.download != null && error?.code === "ENOENT") return null;
    throw error;
  }
}

async function resolveTemplateSkill(templatePath, expectedHash = null) {
  const candidate = path.join(templatePath, "SKILL.md");
  const stat = await fs.lstat(candidate);
  if (!stat.isFile() || stat.isSymbolicLink()) {
    throw new Error("SKILL.md must be a regular nonsymlink file");
  }
  if (stat.size > MAX_SKILL_BYTES) {
    throw new Error(`SKILL.md exceeds the ${MAX_SKILL_BYTES}-byte budget`);
  }
  const [canonicalTemplatePath, canonicalSkillPath] = await Promise.all([
    fs.realpath(templatePath),
    fs.realpath(candidate),
  ]);
  if (!isInside(canonicalTemplatePath, canonicalSkillPath)) {
    throw new Error("SKILL.md escapes the template directory");
  }
  if (expectedHash != null) {
    const actualHash = await sha256File(canonicalSkillPath);
    if (actualHash !== expectedHash) throw new Error("SKILL.md SHA-256 mismatch");
  }
  return canonicalSkillPath;
}

async function assertPresentationTemplateSurface(templatePath, metadata) {
  const allowedRoot = new Set(["SKILL.md", SIDECAR_NAME, "agents", "assets"]);
  const rootEntries = await fs.readdir(templatePath, { withFileTypes: true });
  for (const entry of rootEntries) {
    if (!allowedRoot.has(entry.name) || entry.isSymbolicLink()) {
      throw new Error(`presentation template contains unsupported entry: ${entry.name}`);
    }
  }
  for (const required of allowedRoot) {
    if (!rootEntries.some((entry) => entry.name === required)) {
      throw new Error(`presentation template is missing ${required}`);
    }
  }
  const agentsPath = path.join(templatePath, "agents");
  const agentEntries = await fs.readdir(agentsPath, { withFileTypes: true });
  if (agentEntries.length !== 1 || agentEntries[0].name !== "agent.yaml" ||
      !agentEntries[0].isFile() || agentEntries[0].isSymbolicLink()) {
    throw new Error("presentation template agents must contain only agent.yaml");
  }
  const assetsPath = path.join(templatePath, "assets");
  const assetEntries = await fs.readdir(assetsPath, { withFileTypes: true });
  const hasReferences = metadata.referenceProgram != null || metadata.referencePptx != null;
  const hasLocalReferences = assetEntries.some((entry) => entry.name === "references" && entry.isDirectory());
  const remoteOnlyReferences = hasReferences && !hasLocalReferences &&
    [metadata.referenceProgram, metadata.referencePptx].filter(Boolean).every((reference) => reference.download != null);
  const expectedAssetNames = new Set(["preview.png", "examples"]);
  if (hasLocalReferences || !remoteOnlyReferences) expectedAssetNames.add("references");
  if (assetEntries.length !== expectedAssetNames.size ||
      assetEntries.some((entry) => entry.isSymbolicLink() || !expectedAssetNames.has(entry.name)) ||
      !assetEntries.some((entry) => entry.name === "preview.png" && entry.isFile()) ||
      !assetEntries.some((entry) => entry.name === "examples" && entry.isDirectory())) {
    throw new Error("presentation template assets must match preview, examples, and declared references");
  }
  const examplesPath = path.join(assetsPath, "examples");
  const exampleEntries = await fs.readdir(examplesPath, { withFileTypes: true });
  const expectedFiles = new Set(metadata.examples.map((example) => path.posix.basename(example.path)));
  if (exampleEntries.length !== expectedFiles.size ||
      exampleEntries.some((entry) =>
        !entry.isFile() || entry.isSymbolicLink() || !expectedFiles.has(entry.name))) {
    throw new Error("presentation template examples/ must match metadata exactly");
  }
  if (hasReferences && hasLocalReferences) {
    const referencesPath = path.join(assetsPath, "references");
    const referenceEntries = await listRelativeFiles(referencesPath);
    const expectedReferences = new Set();
    for (const reference of [metadata.referenceProgram, metadata.referencePptx]) {
      if (reference == null) continue;
      const localPath = path.resolve(templatePath, reference.path);
      const relativeToReferences = path.relative(referencesPath, localPath).split(path.sep).join("/");
      const localExists = await fs.lstat(localPath).then((stat) => stat.isFile()).catch(() => false);
      if (localExists) expectedReferences.add(relativeToReferences);
      if (reference === metadata.referenceProgram && localExists) {
        let program;
        try { program = JSON.parse(await fs.readFile(localPath, "utf8")); } catch (error) {
          throw new Error(`referenceProgram is not valid JSON: ${error.message}`);
        }
        for (const declaration of [program.source, ...(program.assets ?? [])]) {
          if (declaration == null || typeof declaration.uri !== "string") continue;
          const dependency = path.resolve(path.dirname(localPath), ...declaration.uri.split("/"));
          const relative = path.relative(referencesPath, dependency).split(path.sep).join("/");
          const dependencyExists = await fs.lstat(dependency).then((stat) => stat.isFile()).catch(() => false);
          if (dependencyExists) expectedReferences.add(relative);
          else if (reference.download == null) throw new Error(`referenceProgram dependency is missing: ${declaration.uri}`);
        }
      }
    }
    if (!referenceEntries.every((entry) => expectedReferences.has(entry)) ||
        !expectedReferences.size && referenceEntries.length > 0 ||
        referenceEntries.length > expectedReferences.size) {
      throw new Error("presentation template references/ must match metadata exactly");
    }
  }
}

async function listRelativeFiles(root, prefix = "") {
  const entries = await fs.readdir(root, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    if (entry.isSymbolicLink()) throw new Error("presentation template references must not contain symlinks");
    const relative = prefix === "" ? entry.name : `${prefix}/${entry.name}`;
    const absolute = path.join(root, entry.name);
    if (entry.isDirectory()) files.push(...await listRelativeFiles(absolute, relative));
    else if (entry.isFile()) files.push(relative);
    else throw new Error("presentation template references contain an unsupported entry");
  }
  return files;
}

function createBm25Context(candidates, intent, tags) {
  const queryTerms = tokenizeValues([
    ...intent.purposes,
    ...intent.audiences,
    ...intent.contentShapes,
    ...intent.visualTraits.tone,
    ...(intent.visualTraits.density == null ? [] : [intent.visualTraits.density]),
    ...(intent.visualTraits.colorMode == null ? [] : [intent.visualTraits.colorMode]),
    ...intent.visualTraits.structure,
    ...tags,
  ]);
  const uniqueQueryTerms = [...new Set(queryTerms)];
  const documents = new Map(
    candidates.map((candidate) => [candidate.id, candidateSearchFields(candidate)]),
  );
  const averageFieldLengths = {};
  for (const field of Object.keys(BM25_FIELD_WEIGHTS)) {
    const total = [...documents.values()].reduce(
      (sum, document) => sum + document[field].length,
      0,
    );
    averageFieldLengths[field] =
      candidates.length === 0 ? 1 : Math.max(total / candidates.length, 1);
  }
  const inverseDocumentFrequency = new Map();
  for (const term of uniqueQueryTerms) {
    let containingDocuments = 0;
    for (const document of documents.values()) {
      if (
        Object.values(document).some((tokens) => tokens.includes(term))
      ) {
        containingDocuments += 1;
      }
    }
    const population = Math.max(candidates.length, 1);
    inverseDocumentFrequency.set(
      term,
      Math.log(
        1 +
          (population - containingDocuments + 0.5) /
            (containingDocuments + 0.5),
      ),
    );
  }
  return {
    averageFieldLengths,
    documents,
    inverseDocumentFrequency,
    queryTerms: uniqueQueryTerms,
  };
}

function candidateSearchFields(candidate) {
  return {
    identity: tokenizeValues([candidate.id, candidate.displayName]),
    useWhen: tokenizeValues(candidate.useWhen),
    audiences: tokenizeValues(candidate.audiences),
    contentShapes: tokenizeValues(candidate.contentShapes),
    tone: tokenizeValues(candidate.visualTraits.tone),
    structure: tokenizeValues(candidate.visualTraits.structure),
    density: tokenizeValues([candidate.visualTraits.density]),
    colorMode: tokenizeValues([candidate.visualTraits.colorMode]),
  };
}

function scoreBm25Candidate(candidate, context) {
  const document = context.documents.get(candidate.id);
  let raw = 0;
  const matchedTerms = [];
  for (const term of context.queryTerms) {
    let weightedFrequency = 0;
    const fields = [];
    for (const [field, weight] of Object.entries(BM25_FIELD_WEIGHTS)) {
      const tokens = document[field];
      const frequency = tokens.filter((token) => token === term).length;
      if (frequency === 0) continue;
      const lengthNormalization =
        1 -
        BM25_B +
        BM25_B *
          (tokens.length / context.averageFieldLengths[field]);
      weightedFrequency += weight * (frequency / lengthNormalization);
      fields.push(field);
    }
    if (weightedFrequency === 0) continue;
    raw +=
      context.inverseDocumentFrequency.get(term) *
      (((BM25_K1 + 1) * weightedFrequency) /
        (BM25_K1 + weightedFrequency));
    matchedTerms.push({ term, fields });
  }
  return {
    raw: roundBm25(raw),
    matchedTerms,
    queryCoverage:
      context.queryTerms.length === 0
        ? 0
        : roundScore(
            (matchedTerms.length / context.queryTerms.length) * 100,
          ),
  };
}

function assessCandidate(candidate, intent, tags, bm25) {
  const matched = [];
  const explainField = (field, queries, values) => {
    for (const query of queries) {
      const best = bestLexicalMatch(query, values);
      if (best.quality >= MIN_FIELD_MATCH) {
        matched.push({
          field,
          query,
          value: best.value,
          quality: roundScore(best.quality * 100),
        });
      }
    }
  };
  explainField(
    "purpose",
    intent.purposes,
    [candidate.id, candidate.displayName, ...candidate.useWhen],
  );
  explainField("audience", intent.audiences, candidate.audiences);
  explainField("contentShape", intent.contentShapes, candidate.contentShapes);
  explainField("tone", intent.visualTraits.tone, candidate.visualTraits.tone);
  explainField(
    "structure",
    intent.visualTraits.structure,
    candidate.visualTraits.structure,
  );
  explainField(
    "density",
    intent.visualTraits.density == null ? [] : [intent.visualTraits.density],
    [candidate.visualTraits.density],
  );
  explainField(
    "colorMode",
    intent.visualTraits.colorMode == null ? [] : [intent.visualTraits.colorMode],
    [candidate.visualTraits.colorMode],
  );

  const positiveFields = Object.values(candidateSearchFields(candidate))
    .flat();
  const matchedTags = [];
  for (const tag of tags) {
    const best = bestLexicalMatch(tag, positiveFields);
    if (best.quality >= MIN_FIELD_MATCH) {
      matchedTags.push(tag);
      matched.push({
        field: "legacyTag",
        query: tag,
        value: best.value,
        quality: roundScore(best.quality * 100),
      });
    }
  }

  const conflictSignals = [
    ...intent.purposes,
    ...intent.audiences,
    ...intent.contentShapes,
    ...intent.visualTraits.tone,
    ...(intent.visualTraits.density == null ? [] : [intent.visualTraits.density]),
    ...(intent.visualTraits.colorMode == null ? [] : [intent.visualTraits.colorMode]),
    ...intent.visualTraits.structure,
    ...tags,
  ];
  const conflicts = [];
  for (const avoid of candidate.avoidWhen) {
    const best = bestLexicalMatch(avoid, conflictSignals);
    if (best.quality >= AVOID_CONFLICT_MATCH) {
      conflicts.push({
        avoidWhen: avoid,
        query: best.value,
        quality: roundScore(best.quality * 100),
      });
    }
  }

  const verifiedOperations = new Set(
    (candidate.editProfile?.verifiedOperations ?? []).map(normalizeTag),
  );
  const missingOperations = intent.requiredOperations.filter(
    (operation) => !verifiedOperations.has(operation),
  );
  const bm25Score = scoreBm25Candidate(candidate, bm25);
  const reviewFlags = [];
  if (intent.brandSensitive) reviewFlags.push("brand-sensitive");
  if (candidate.visualCommitment === "opinionated") reviewFlags.push("opinionated-template");
  const rejectionReasons = [];
  if (conflicts.length > 0) rejectionReasons.push("avoid-when-conflict");
  if (missingOperations.length > 0) rejectionReasons.push("missing-verified-operation");
  if (bm25.queryTerms.length > 0 && bm25Score.raw === 0) {
    rejectionReasons.push("insufficient-relevance");
  }

  return {
    matchedTags,
    match: {
      score: 0,
      bm25: bm25Score.raw,
      queryCoverage: bm25Score.queryCoverage,
      matchedTerms: bm25Score.matchedTerms,
      matched,
      conflicts,
      missingOperations,
    },
    reviewFlags,
    rejectionReasons,
  };
}

function bestLexicalMatch(query, values) {
  let best = { quality: 0, value: null };
  for (const value of values) {
    const quality = lexicalSimilarity(query, value);
    if (quality > best.quality) best = { quality, value };
  }
  return best;
}

function lexicalSimilarity(left, right) {
  const query = normalizeSearchText(left);
  const candidate = normalizeSearchText(right);
  if (query.length === 0 || candidate.length === 0) return 0;
  if (query === candidate) return 1;
  const queryTokens = new Set(query.split(" "));
  const candidateTokens = new Set(candidate.split(" "));
  let overlap = 0;
  for (const token of queryTokens) {
    if (candidateTokens.has(token)) overlap += 1;
  }
  if (overlap === 0) return 0;
  const coverage = overlap / queryTokens.size;
  const precision = overlap / candidateTokens.size;
  const tokenScore = 0.7 * coverage + 0.3 * precision;
  if (candidate.includes(query) || query.includes(candidate)) {
    const lengthRatio =
      Math.min(queryTokens.size, candidateTokens.size) /
      Math.max(queryTokens.size, candidateTokens.size);
    return Math.max(tokenScore, 0.85 + 0.15 * lengthRatio);
  }
  return tokenScore;
}

function tokenizeValues(values) {
  return values.flatMap((value) => {
    const normalized = normalizeSearchText(value);
    return normalized.length === 0 ? [] : normalized.split(" ");
  });
}

function normalizeIntent(intent) {
  const result = {
    purposes: [],
    audiences: [],
    contentShapes: [],
    visualTraits: {
      tone: [],
      density: null,
      colorMode: null,
      structure: [],
    },
    requiredOperations: [],
    brandSensitive: false,
  };
  if (intent == null) return result;
  if (typeof intent !== "object" || Array.isArray(intent)) {
    throw new Error("intent must be an object.");
  }
  assertObjectKeys(
    intent,
    "intent",
    [
      "purposes",
      "audiences",
      "contentShapes",
      "visualTraits",
      "requiredOperations",
      "brandSensitive",
    ],
  );
  result.purposes = normalizeIntentValues(intent.purposes, "intent.purposes");
  result.audiences = normalizeIntentValues(intent.audiences, "intent.audiences");
  result.contentShapes = normalizeIntentValues(
    intent.contentShapes,
    "intent.contentShapes",
  );
  result.requiredOperations = normalizeIntentValues(
    intent.requiredOperations,
    "intent.requiredOperations",
    normalizeTag,
  );
  if (intent.visualTraits != null) {
    if (
      typeof intent.visualTraits !== "object" ||
      Array.isArray(intent.visualTraits)
    ) {
      throw new Error("intent.visualTraits must be an object.");
    }
    assertObjectKeys(
      intent.visualTraits,
      "intent.visualTraits",
      ["tone", "density", "colorMode", "structure"],
    );
    result.visualTraits.tone = normalizeIntentValues(
      intent.visualTraits.tone,
      "intent.visualTraits.tone",
    );
    result.visualTraits.structure = normalizeIntentValues(
      intent.visualTraits.structure,
      "intent.visualTraits.structure",
    );
    if (intent.visualTraits.density != null) {
      assertEnum(intent.visualTraits.density, "intent.visualTraits.density", VALID_DENSITIES);
      result.visualTraits.density = intent.visualTraits.density;
    }
    if (intent.visualTraits.colorMode != null) {
      assertEnum(
        intent.visualTraits.colorMode,
        "intent.visualTraits.colorMode",
        VALID_COLOR_MODES,
      );
      result.visualTraits.colorMode = intent.visualTraits.colorMode;
    }
  }
  if (
    intent.brandSensitive != null &&
    typeof intent.brandSensitive !== "boolean"
  ) {
    throw new Error("intent.brandSensitive must be a boolean.");
  }
  result.brandSensitive = intent.brandSensitive ?? false;
  return result;
}

function normalizeIntentValues(value, label, normalizer = normalizeSearchText) {
  if (value == null) return [];
  if (!Array.isArray(value) || value.length > MAX_INTENT_VALUES) {
    throw new Error(`${label} must be an array of at most ${MAX_INTENT_VALUES} strings.`);
  }
  const normalized = value.map((entry) => {
    if (
      typeof entry !== "string" ||
      entry.trim().length === 0 ||
      entry.length > 120 ||
      /[\0\r\n]/u.test(entry)
    ) {
      throw new Error(`${label} entries must be one non-empty line of at most 120 characters.`);
    }
    return normalizer(entry);
  });
  if (normalized.some((entry) => entry.length === 0)) {
    throw new Error(`${label} entries must contain searchable text.`);
  }
  return [...new Set(normalized)];
}

function normalizeQueryTags(tags) {
  if (!Array.isArray(tags)) throw new Error("tags must be an array.");
  if (tags.length > 20) throw new Error("At most 20 query tags may be used.");
  const normalized = tags.map((tag) => {
    if (typeof tag !== "string" || tag.trim().length === 0 || tag.length > 80) {
      throw new Error("Each query tag must be a non-empty string of at most 80 characters.");
    }
    return normalizeTag(tag);
  });
  return [...new Set(normalized)];
}

function normalizeTag(value) {
  return normalizeSearchText(value)
    .replace(/\s+/gu, "-");
}

function normalizeSearchText(value) {
  return value
    .normalize("NFKC")
    .toLowerCase()
    .trim()
    .replace(/[^\p{Letter}\p{Number}]+/gu, " ")
    .trim()
    .replace(/\s+/gu, " ");
}

function roundScore(value) {
  return Math.round(value * 10) / 10;
}

function roundBm25(value) {
  return Math.round(value * 10_000) / 10_000;
}

function assertTemplateId(value, label, optional = false) {
  if (optional && value == null) return;
  if (typeof value !== "string" || !TEMPLATE_NAME_PATTERN.test(value)) {
    throw new Error(`${label} must be an artifact-template-* identifier.`);
  }
}

function assertKind(value) {
  if (!VALID_KINDS.has(value)) {
    throw new Error("kind must be document, spreadsheet, or presentation.");
  }
}

function assertShortString(value, label, max) {
  if (
    typeof value !== "string" ||
    value.trim().length === 0 ||
    value !== value.trim() ||
    value.length > max ||
    /[\0\r\n]/u.test(value)
  ) {
    throw new Error(`${label} must be one trimmed line of at most ${max} characters`);
  }
}

function assertStringArray(value, label, { min, max }) {
  if (!Array.isArray(value) || value.length < min || value.length > max) {
    throw new Error(`${label} must contain ${min}-${max} strings`);
  }
  const seen = new Set();
  for (const entry of value) {
    assertShortString(entry, label, 120);
    const normalized = normalizeTag(entry);
    if (seen.has(normalized)) throw new Error(`${label} must not contain duplicates`);
    seen.add(normalized);
  }
}

function assertEnglishSearchArray(value, label, bounds) {
  assertStringArray(value, label, bounds);
  for (const entry of value) {
    if (!/^[\x20-\x7e]+$/u.test(entry) || !/[a-z]/iu.test(entry)) {
      throw new Error(`${label} must use English search text`);
    }
  }
}

function assertEnum(value, label, allowed) {
  if (!allowed.has(value)) {
    throw new Error(`${label} must be one of ${[...allowed].join(", ")}`);
  }
}

function assertAssetId(value, label) {
  if (typeof value !== "string" || value.trim() !== value || value.length === 0 || value.length > 512 || /[\0\r\n]/u.test(value)) {
    throw new Error(`${label} must be a bounded non-empty asset identifier`);
  }
}

function validateReplacementDimension(value, label) {
  if (!Number.isSafeInteger(value) || value < 1 || value > 16_384) {
    throw new Error(`${label} must be an integer from 1 to 16384`);
  }
}

function assertObjectKeys(value, label, allowedKeys) {
  const allowed = new Set(allowedKeys);
  const extra = Object.keys(value).filter((key) => !allowed.has(key));
  if (extra.length > 0) {
    throw new Error(`${label} contains unsupported fields: ${extra.join(", ")}`);
  }
}

function assertHash(value, label) {
  if (typeof value !== "string" || !HASH_PATTERN.test(value)) {
    throw new Error(`${label} must be a lowercase SHA-256 value`);
  }
}

function assertRelativeAssetPath(value, label) {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    path.isAbsolute(value) ||
    value.includes("\\") ||
    value.split("/").some((segment) => segment === "" || segment === "." || segment === "..")
  ) {
    throw new Error(`${label} must be a safe relative path`);
  }
}

function assertRelativePngPath(value, label) {
  assertRelativeAssetPath(value, label);
  if (path.posix.extname(value).toLowerCase() !== ".png") {
    throw new Error(`${label} must use a .png file`);
  }
}

function isInside(root, candidate) {
  const relative = path.relative(root, candidate);
  return relative.length > 0 && relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

function isInsideOrEqual(root, candidate) {
  const relative = path.relative(root, candidate);
  return relative === "" ||
    (relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative));
}

async function sha256File(filePath) {
  const hash = createHash("sha256");
  for await (const chunk of createReadStream(filePath)) hash.update(chunk);
  return hash.digest("hex");
}

export function parseTemplateSearchArguments(args) {
  const request = {
    tags: [],
    roots: [],
    intent: {
      purposes: [],
      audiences: [],
      contentShapes: [],
      visualTraits: {
        tone: [],
        structure: [],
      },
      requiredOperations: [],
      brandSensitive: false,
    },
  };
  for (let index = 0; index < args.length; index += 1) {
    const flag = args[index];
    if (flag === "--help" || flag === "-h") {
      return { help: true };
    }
    if (flag === "--json") {
      request.json = true;
      continue;
    }
    if (flag === "--brand-sensitive") {
      request.intent.brandSensitive = true;
      continue;
    }
    const value = args[index + 1];
    if (value == null || value.startsWith("--")) throw new Error(TEMPLATE_SEARCH_USAGE);
    index += 1;
    if (flag === "--kind") request.kind = value;
    else if (flag === "--tag") request.tags.push(value);
    else if (flag === "--purpose") request.intent.purposes.push(value);
    else if (flag === "--audience") request.intent.audiences.push(value);
    else if (flag === "--content-shape") request.intent.contentShapes.push(value);
    else if (flag === "--tone") request.intent.visualTraits.tone.push(value);
    else if (flag === "--structure") request.intent.visualTraits.structure.push(value);
    else if (flag === "--density") request.intent.visualTraits.density = value;
    else if (flag === "--color-mode") request.intent.visualTraits.colorMode = value;
    else if (flag === "--operation") request.intent.requiredOperations.push(value);
    else if (flag === "--id") request.id = value;
    else if (flag === "--root") request.roots.push(value);
    else if (flag === "--max") request.maxCandidates = Number(value);
    else throw new Error(TEMPLATE_SEARCH_USAGE);
  }
  if (request.kind == null) throw new Error(TEMPLATE_SEARCH_USAGE);
  if (request.roots.length === 0) request.roots = null;
  return request;
}

export function formatTemplateSearchResult(result) {
  if (result.candidates.length === 0) {
    return [
      `No ${result.kind} template matched the search.`,
      `Searched ${result.searchedRoots.length} catalog root${result.searchedRoots.length === 1 ? "" : "s"}; ` +
        `${result.rejected.length} rejected, ${result.invalid.length} invalid.`,
      "Template decision: none remains available.",
    ].join("\n");
  }
  const rows = [
    ["Rank", "Template", "Score", "Coverage", "Review"],
    ...result.candidates.map((candidate, index) => [
      String(index + 1),
      `${candidate.displayName} (${candidate.id})`,
      candidate.match.score.toFixed(1),
      `${candidate.match.queryCoverage.toFixed(1)}%`,
      candidate.reviewFlags.length === 0 ? "-" : candidate.reviewFlags.join(","),
    ]),
  ];
  const widths = rows[0].map((_, column) =>
    Math.max(...rows.map((row) => row[column].length)));
  return [
    ...rows.map((row) =>
      row.map((value, column) => value.padEnd(widths[column])).join("  ").trimEnd()),
    "",
    `Returned ${result.candidates.length} candidate${result.candidates.length === 1 ? "" : "s"}; ` +
      `${result.rejected.length} rejected, ${result.invalid.length} invalid.`,
    "Selection remains with the Agent (selected, ask, or none).",
  ].join("\n");
}
