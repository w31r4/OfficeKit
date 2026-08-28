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
const DEFAULT_MAX_CANDIDATES = 5;
const MAX_CANDIDATES = 20;
const MAX_INTENT_VALUES = 20;
const MIN_FIELD_MATCH = 0.45;
const AVOID_CONFLICT_MATCH = 0.72;
const BM25_K1 = 1.2;
const BM25_B = 0.75;
const VALID_KINDS = new Set(["document", "presentation", "spreadsheet"]);
const PRESENTATION_TEMPLATE_SCHEMA_VERSION = 3;
const REFERENCE_EXTENSIONS = new Map([
  ["document", ".docx"],
  ["presentation", ".pptx"],
  ["spreadsheet", ".xlsx"],
]);
const VALID_DENSITIES = new Set(["sparse", "medium", "dense", "mixed"]);
const VALID_COLOR_MODES = new Set(["light", "dark", "neutral", "mixed"]);
const VALID_COMMITMENTS = new Set(["neutral", "opinionated"]);
const VALID_EDIT_LEVELS = new Set(["copy-only", "bounded-edit", "composable"]);
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

  const skillPath = await resolveTemplateSkill(templatePath);
  const previewPath = await resolveAsset(
    templatePath,
    metadata.preview,
    metadata.provenance.previewSha256,
    "preview",
  );
  const referencePath = metadata.schemaVersion === PRESENTATION_TEMPLATE_SCHEMA_VERSION
    ? undefined
    : await resolveAsset(
      templatePath,
      metadata.reference,
      metadata.provenance.referenceSha256,
      "reference",
    );
  const examples = metadata.schemaVersion === PRESENTATION_TEMPLATE_SCHEMA_VERSION
    ? await Promise.all(metadata.examples.map(async (example) => ({
      path: example.path,
      role: example.role,
      sha256: example.sha256,
      absolutePath: await resolveAsset(templatePath, example.path, example.sha256, `example ${example.path}`),
    })))
    : [];

  return {
    id: metadata.id,
    displayName: metadata.displayName,
    kind: metadata.kind,
    useWhen: metadata.useWhen,
    avoidWhen: metadata.avoidWhen,
    audiences: metadata.audiences,
    contentShapes: metadata.contentShapes,
    visualTraits: metadata.visualTraits,
    visualCommitment: metadata.visualCommitment,
    editProfile: metadata.editProfile,
    templateSchemaVersion: metadata.schemaVersion,
    examples,
    provenance: {
      license: metadata.provenance.license,
      source: metadata.provenance.source,
      referenceSha256: metadata.provenance.referenceSha256,
      previewSha256: metadata.provenance.previewSha256,
    },
    catalogSource: root.source,
    templateRoot: await fs.realpath(templatePath),
    skillPath,
    referencePath,
    previewPath,
  };
}

function validateMetadata(value, expectedId) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("metadata must be an object");
  }
  const isCleanRoomPresentation = value.schemaVersion === PRESENTATION_TEMPLATE_SCHEMA_VERSION && value.kind === "presentation";
  assertObjectKeys(value, "metadata", isCleanRoomPresentation
    ? ["schemaVersion", "id", "displayName", "kind", "preview", "examples", "useWhen", "avoidWhen", "audiences", "contentShapes", "visualTraits", "visualCommitment", "editProfile", "provenance"]
    : ["schemaVersion", "id", "displayName", "kind", "reference", "preview", "useWhen", "avoidWhen", "audiences", "contentShapes", "visualTraits", "visualCommitment", "editProfile", "provenance"]);
  if (value.schemaVersion !== 2 && !isCleanRoomPresentation) {
    throw new Error("schemaVersion must be 2 for legacy templates or 3 for clean-room presentations");
  }
  assertTemplateId(value.id, "id");
  if (value.id !== expectedId) {
    throw new Error(`id must match directory name ${expectedId}`);
  }
  assertShortString(value.displayName, "displayName", 80);
  assertKind(value.kind);
  assertEnglishSearchArray(value.useWhen, "useWhen", { min: 1, max: 20 });
  assertEnglishSearchArray(value.avoidWhen, "avoidWhen", { min: 0, max: 20 });
  assertEnglishSearchArray(value.audiences, "audiences", { min: 0, max: 20 });
  assertEnglishSearchArray(value.contentShapes, "contentShapes", { min: 0, max: 20 });
  assertRelativeAssetPath(value.preview, "preview");
  if (path.posix.extname(value.preview).toLowerCase() !== ".png") {
    throw new Error("preview must use a .png file");
  }
  if (isCleanRoomPresentation) {
    if (!Array.isArray(value.examples) || value.examples.length < 1 || value.examples.length > 8) {
      throw new Error("presentation schemaVersion 3 templates must declare 1-8 examples");
    }
    const paths = new Set();
    for (const [index, example] of value.examples.entries()) {
      if (example == null || typeof example !== "object" || Array.isArray(example)) throw new Error(`examples[${index}] must be an object`);
      assertObjectKeys(example, `examples[${index}]`, ["path", "role", "sha256"]);
      assertRelativeAssetPath(example.path, `examples[${index}].path`);
      if (path.posix.extname(example.path).toLowerCase() !== ".png") throw new Error(`examples[${index}].path must use a .png file`);
      assertShortString(example.role, `examples[${index}].role`, 80);
      assertHash(example.sha256, `examples[${index}].sha256`);
      if (paths.has(example.path)) throw new Error("examples must not repeat an asset path");
      paths.add(example.path);
    }
  } else {
    assertRelativeAssetPath(value.reference, "reference");
    if (path.posix.extname(value.reference).toLowerCase() !== REFERENCE_EXTENSIONS.get(value.kind)) {
      throw new Error(`${value.kind} templates must use a ${REFERENCE_EXTENSIONS.get(value.kind)} reference`);
    }
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
  assertObjectKeys(value.provenance, "provenance", isCleanRoomPresentation
    ? ["license", "source", "previewSha256"]
    : ["license", "source", "referenceSha256", "previewSha256"]);
  assertShortString(value.provenance.license, "provenance.license", 120);
  assertShortString(value.provenance.source, "provenance.source", 500);
  if (!isCleanRoomPresentation) assertHash(value.provenance.referenceSha256, "provenance.referenceSha256");
  assertHash(value.provenance.previewSha256, "provenance.previewSha256");
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

async function resolveTemplateSkill(templatePath) {
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
  return canonicalSkillPath;
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
    candidate.editProfile.verifiedOperations.map(normalizeTag),
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
