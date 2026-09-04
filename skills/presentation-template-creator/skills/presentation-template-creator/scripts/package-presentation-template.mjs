#!/usr/bin/env node

import { createHash, randomUUID } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { deflateSync, inflateSync } from "node:zlib";

const TEMPLATE_ID = /^artifact-template-[a-z0-9]+(?:-[a-z0-9]+)*$/u;
const HASH = /^[a-f0-9]{64}$/u;
const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const VALID_DENSITIES = new Set(["sparse", "medium", "dense", "mixed"]);
const VALID_COLOR_MODES = new Set(["light", "dark", "neutral", "mixed"]);
const VALID_COMMITMENTS = new Set(["neutral", "opinionated"]);
const VALID_ROLES = new Set([
  "cover", "section", "analysis", "data", "process", "comparison", "closing", "mixed",
]);
const VALID_IMAGE_SLOT_ROLES = new Set([
  "hero", "thumbnail", "avatar", "background", "logo", "diagram", "screenshot", "photo", "icon", "chart-source", "any",
]);
const VALID_IMAGE_SLOT_FITS = new Set(["contain", "cover", "stretch"]);
const VALID_IMAGE_SLOT_MASKS = new Set(["none", "rect", "roundRect", "ellipse", "custom"]);
const VALID_IMAGE_SLOT_RIGHTS = new Set([
  "user-provided", "generated", "permission", "public-domain", "cc0", "cc-by", "official-press-kit", "internal", "other",
]);
const IMAGE_SLOT_ID_PATTERN = /^[a-z][a-z0-9-]{0,63}$/u;
const MAX_SPEC_BYTES = 256 * 1024;
const MAX_GUIDE_BYTES = 256 * 1024;
const MAX_IMAGE_BYTES = 20 * 1024 * 1024;
const MAX_TOTAL_IMAGE_BYTES = 60 * 1024 * 1024;
const MAX_IMAGE_PIXELS = 40_000_000;
const MAX_REFERENCE_PROGRAM_BYTES = 16 * 1024 * 1024;
const MAX_REFERENCE_PPTX_BYTES = 256 * 1024 * 1024;
const MAX_REFERENCE_ASSET_BYTES = 64 * 1024 * 1024;
const MAX_TOTAL_REFERENCE_ASSET_BYTES = 256 * 1024 * 1024;
const REMOTE_REFERENCE_HOSTS = new Set(["raw.githubusercontent.com"]);
const LOCK_NAME = ".presentation-template-write-lock";

let args = null;
try {
  args = parseArguments(process.argv.slice(2));
  const result = await packageTemplate(args);
  if (args.json) process.stdout.write(`${JSON.stringify({ ok: true, ...result })}\n`);
  else process.stdout.write(`Created ${result.skillName} at ${result.skillPath}\n`);
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  if (args?.json || process.argv.includes("--json")) process.stderr.write(`${JSON.stringify({ ok: false, error: message })}\n`);
  else process.stderr.write(`Presentation template packaging failed: ${message}\n`);
  process.exitCode = 2;
}

async function packageTemplate({ specPath, outputRoot, expectedSha256 }) {
  const spec = await readJsonFile(specPath, MAX_SPEC_BYTES, "spec");
  validateSpec(spec);
  const guideBody = await readTextFile(spec.guidePath, MAX_GUIDE_BYTES, "guidePath");
  if (guideBody.trim().length < 120) throw new Error("guidePath must contain a substantive style guide");
  if (/^---\s*$/mu.test(guideBody)) throw new Error("guidePath must be a Markdown body without YAML frontmatter");
  const referenceProgram = await readReference(spec.referenceProgram, {
    label: "referenceProgram",
    maxBytes: MAX_REFERENCE_PROGRAM_BYTES,
    extension: ".ppj",
  });
  let referenceProgramAssets = [];
  if (referenceProgram != null) {
    let parsed;
    try {
      parsed = JSON.parse(referenceProgram.bytes.toString("utf8"));
    } catch (error) {
      throw new Error(`referenceProgram is not strict JSON: ${error.message}`);
    }
    if (parsed?.schema !== "office-kit/ppj/v1") {
      throw new Error("referenceProgram must use schema office-kit/ppj/v1");
    }
    referenceProgramAssets = await readReferenceProgramAssets(parsed, spec.referenceProgram.path);
  }
  const referencePptx = await readReference(spec.referencePptx, {
    label: "referencePptx",
    maxBytes: MAX_REFERENCE_PPTX_BYTES,
    extension: ".pptx",
  });
  if (referencePptx != null && !referencePptx.bytes.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 0x03, 0x04]))) {
    throw new Error("referencePptx must be an OPC ZIP package");
  }

  const examples = [];
  let totalImageBytes = 0;
  for (const [index, example] of spec.examples.entries()) {
    const bytes = await readRegularFile(example.path, MAX_IMAGE_BYTES, `examples[${index}].path`);
    totalImageBytes += bytes.byteLength;
    if (totalImageBytes > MAX_TOTAL_IMAGE_BYTES) {
      throw new Error(`example images exceed the ${MAX_TOTAL_IMAGE_BYTES}-byte total budget`);
    }
    const decoded = decodePngRgba(bytes);
    if (decoded.width * decoded.height > MAX_IMAGE_PIXELS) {
      throw new Error(`examples[${index}] exceeds the ${MAX_IMAGE_PIXELS}-pixel budget`);
    }
    if (decoded.width < 320 || decoded.height < 180 || decoded.width > 8192 || decoded.height > 8192) {
      throw new Error(`examples[${index}] dimensions must be between 320x180 and 8192x8192`);
    }
    examples.push({ ...example, bytes, decoded });
  }

  const absoluteOutputRoot = path.resolve(outputRoot);
  await fs.mkdir(absoluteOutputRoot, { recursive: true, mode: 0o755 });
  const outputStat = await fs.lstat(absoluteOutputRoot);
  if (!outputStat.isDirectory() || outputStat.isSymbolicLink()) {
    throw new Error("output root must be a real directory");
  }
  const canonicalOutputRoot = await fs.realpath(absoluteOutputRoot);
  const targetPath = path.join(canonicalOutputRoot, spec.id);
  const lockPath = path.join(canonicalOutputRoot, LOCK_NAME);
  await acquireLock(lockPath);
  const stagedPath = path.join(canonicalOutputRoot, `.${spec.id}.stage-${randomUUID()}`);
  let stagedExists = false;
  try {
    const current = await inspectCurrentTarget(targetPath);
    if (current == null && expectedSha256 != null) {
      throw new Error("--expected-sha256 was provided but the target does not exist");
    }
    if (current != null) {
      if (expectedSha256 == null) {
        throw new Error("target already exists; pass --expected-sha256 for an explicit update");
      }
      if (current.sidecarSha256 !== expectedSha256) {
        throw new Error("target changed since inspection; expected sidecar SHA-256 does not match");
      }
    }

    await fs.mkdir(path.join(stagedPath, "agents"), { recursive: true, mode: 0o755 });
    await fs.mkdir(path.join(stagedPath, "assets", "examples"), { recursive: true, mode: 0o755 });
    if (referenceProgram != null || referencePptx != null) {
      await fs.mkdir(path.join(stagedPath, "assets", "references"), { recursive: true, mode: 0o755 });
    }
    stagedExists = true;

    const skillText = `---\nname: ${spec.id}\ndescription: ${JSON.stringify(spec.description)}\n---\n\n${guideBody.trim()}\n`;
    const stagedExamples = [];
    for (const [index, example] of examples.entries()) {
      const fileName = `${String(index + 1).padStart(2, "0")}-${example.role}.png`;
      const relativePath = `assets/examples/${fileName}`;
      await writeImmutable(path.join(stagedPath, relativePath), example.bytes);
      stagedExamples.push({
        path: relativePath,
        role: example.role,
        sha256: sha256(example.bytes),
      });
    }
    const packagedImageSlots = spec.imageSlots == null ? null : spec.imageSlots.map((slot) => ({
      ...slot,
      examplePath: stagedExamples[examples.findIndex((example) =>
        path.resolve(example.path) === path.resolve(slot.examplePath))].path,
    }));
    const previewBytes = encodePngRgba(createMontage(examples.map((example) => example.decoded)));
    const agentText = [
      "interface:",
      `  display_name: ${JSON.stringify(spec.displayName)}`,
      `  short_description: ${JSON.stringify(`Use the ${spec.displayName} presentation style`)}`,
      '  icon_large: "./assets/preview.png"',
      `  default_prompt: ${JSON.stringify(`Use $${spec.id} as the visual style for this presentation.`)}`,
      "",
    ].join("\n");
    await Promise.all([
      writeImmutable(path.join(stagedPath, "SKILL.md"), skillText),
      writeImmutable(path.join(stagedPath, "agents", "agent.yaml"), agentText),
      writeImmutable(path.join(stagedPath, "assets", "preview.png"), previewBytes),
      ...(referenceProgram == null ? [] : [
        writeImmutable(path.join(stagedPath, "assets", "references", "reference.ppj"), referenceProgram.bytes),
      ]),
      ...(referencePptx == null ? [] : [
        writeImmutable(path.join(stagedPath, "assets", "references", "reference.pptx"), referencePptx.bytes),
      ]),
    ]);
    for (const asset of referenceProgramAssets) {
      const assetPath = path.join(stagedPath, "assets", "references", asset.relativePath);
      await fs.mkdir(path.dirname(assetPath), { recursive: true, mode: 0o755 });
      await writeImmutable(assetPath, asset.bytes);
    }

    const sidecar = {
      schemaVersion: 3,
      id: spec.id,
      displayName: spec.displayName,
      kind: "presentation",
      preview: "assets/preview.png",
      examples: stagedExamples,
      useWhen: spec.useWhen,
      avoidWhen: spec.avoidWhen,
      audiences: spec.audiences,
      contentShapes: spec.contentShapes,
      visualTraits: spec.visualTraits,
      visualCommitment: spec.visualCommitment,
      ...(packagedImageSlots == null ? {} : { imageSlots: packagedImageSlots }),
      ...(referenceProgram == null ? {} : {
        referenceProgram: referenceMetadata(referenceProgram, "assets/references/reference.ppj"),
      }),
      ...(referencePptx == null ? {} : {
        referencePptx: referenceMetadata(referencePptx, "assets/references/reference.pptx"),
      }),
      provenance: {
        license: spec.provenance.license,
        source: spec.provenance.source,
        guideSha256: sha256(skillText),
        previewSha256: sha256(previewBytes),
      },
    };
    const sidecarBytes = Buffer.from(`${JSON.stringify(sidecar, null, 2)}\n`);
    await writeImmutable(path.join(stagedPath, "artifact-template.json"), sidecarBytes);
    await validateStagedSurface(stagedPath, sidecar, referenceProgramAssets);
    await publishAtomically({ targetPath, stagedPath, current });
    stagedExists = false;
    return {
      schemaVersion: 3,
      skillName: spec.id,
      skillPath: targetPath,
      sidecarSha256: sha256(sidecarBytes),
      previewPath: path.join(targetPath, "assets", "preview.png"),
      examplePaths: stagedExamples.map((entry) => path.join(targetPath, entry.path)),
      referenceProgramPath: referenceProgram == null ? null : path.join(targetPath, "assets", "references", "reference.ppj"),
      referencePptxPath: referencePptx == null ? null : path.join(targetPath, "assets", "references", "reference.pptx"),
      updated: current != null,
    };
  } finally {
    if (stagedExists) await fs.rm(stagedPath, { recursive: true, force: true });
    await fs.rm(lockPath, { recursive: true, force: true });
  }
}

function validateSpec(spec) {
  if (spec == null || typeof spec !== "object" || Array.isArray(spec)) throw new Error("spec must be an object");
  assertKeysWithOptional(spec, "spec", [
    "id", "displayName", "description", "guidePath", "useWhen", "avoidWhen",
    "audiences", "contentShapes", "visualTraits", "visualCommitment", "examples",
    "referenceProgram", "referencePptx", "provenance",
  ], ["imageSlots"]);
  if (!TEMPLATE_ID.test(spec.id ?? "")) throw new Error("id must be an artifact-template-* identifier");
  assertLine(spec.displayName, "displayName", 80);
  assertLine(spec.description, "description", 320);
  assertAbsolutePath(spec.guidePath, "guidePath");
  assertEnglishArray(spec.useWhen, "useWhen", 1, 20);
  assertEnglishArray(spec.avoidWhen, "avoidWhen", 0, 20);
  assertEnglishArray(spec.audiences, "audiences", 0, 20);
  assertEnglishArray(spec.contentShapes, "contentShapes", 0, 20);
  if (spec.visualTraits == null || typeof spec.visualTraits !== "object" || Array.isArray(spec.visualTraits)) {
    throw new Error("visualTraits must be an object");
  }
  assertKeys(spec.visualTraits, "visualTraits", ["tone", "density", "colorMode", "structure"]);
  assertEnglishArray(spec.visualTraits.tone, "visualTraits.tone", 0, 12);
  assertEnglishArray(spec.visualTraits.structure, "visualTraits.structure", 0, 12);
  assertEnum(spec.visualTraits.density, "visualTraits.density", VALID_DENSITIES);
  assertEnum(spec.visualTraits.colorMode, "visualTraits.colorMode", VALID_COLOR_MODES);
  assertEnum(spec.visualCommitment, "visualCommitment", VALID_COMMITMENTS);
  if (!Array.isArray(spec.examples) || spec.examples.length < 4 || spec.examples.length > 6) {
    throw new Error("examples must contain 4-6 entries");
  }
  const roles = new Set();
  const paths = new Set();
  for (const [index, example] of spec.examples.entries()) {
    if (example == null || typeof example !== "object" || Array.isArray(example)) {
      throw new Error(`examples[${index}] must be an object`);
    }
    assertKeys(example, `examples[${index}]`, ["path", "role"]);
    assertAbsolutePath(example.path, `examples[${index}].path`);
    assertEnum(example.role, `examples[${index}].role`, VALID_ROLES);
    if (paths.has(path.resolve(example.path))) throw new Error("example paths must be unique");
    paths.add(path.resolve(example.path));
    roles.add(example.role);
  }
  if (roles.size < 3) throw new Error("examples must cover at least 3 distinct roles");
  validatePresentationImageSlots(spec.imageSlots, spec.examples);
  validateReferenceSpec(spec.referenceProgram, "referenceProgram", ".ppj");
  validateReferenceSpec(spec.referencePptx, "referencePptx", ".pptx");
  if (spec.provenance == null || typeof spec.provenance !== "object" || Array.isArray(spec.provenance)) {
    throw new Error("provenance must be an object");
  }
  assertKeys(spec.provenance, "provenance", ["license", "source"]);
  assertLine(spec.provenance.license, "provenance.license", 120);
  assertLine(spec.provenance.source, "provenance.source", 500);
}

function validatePresentationImageSlots(value, examples) {
  if (value == null) return;
  if (!Array.isArray(value) || value.length > 64) {
    throw new Error("imageSlots must contain 0-64 entries");
  }
  const examplePaths = new Set(examples.map((example) => path.resolve(example.path)));
  const ids = new Set();
  for (const [index, slot] of value.entries()) {
    if (slot == null || typeof slot !== "object" || Array.isArray(slot)) {
      throw new Error(`imageSlots[${index}] must be an object`);
    }
    assertOptionalKeys(slot, `imageSlots[${index}]`, [
      "id", "role", "examplePath", "allowedFit", "allowedMask", "minWidthPx", "minHeightPx", "rights",
    ]);
    if (typeof slot.id !== "string" || !IMAGE_SLOT_ID_PATTERN.test(slot.id)) {
      throw new Error(`imageSlots[${index}].id must be a lowercase identifier`);
    }
    if (ids.has(slot.id)) throw new Error(`imageSlots must use unique ids: ${slot.id}`);
    ids.add(slot.id);
    assertEnum(slot.role, `imageSlots[${index}].role`, VALID_IMAGE_SLOT_ROLES);
    assertAbsolutePath(slot.examplePath, `imageSlots[${index}].examplePath`);
    if (!examplePaths.has(path.resolve(slot.examplePath))) {
      throw new Error(`imageSlots[${index}].examplePath must reference a declared example`);
    }
    validateBoundedEnumArray(slot.allowedFit, `imageSlots[${index}].allowedFit`, VALID_IMAGE_SLOT_FITS, 3);
    validateBoundedEnumArray(slot.allowedMask, `imageSlots[${index}].allowedMask`, VALID_IMAGE_SLOT_MASKS, 16);
    validateBoundedEnumArray(slot.rights, `imageSlots[${index}].rights`, VALID_IMAGE_SLOT_RIGHTS, 16);
    validateOptionalPixelDimension(slot.minWidthPx, `imageSlots[${index}].minWidthPx`);
    validateOptionalPixelDimension(slot.minHeightPx, `imageSlots[${index}].minHeightPx`);
  }
}

function validateBoundedEnumArray(value, label, allowed, max) {
  if (value == null) return;
  if (!Array.isArray(value) || value.length < 1 || value.length > max) {
    throw new Error(`${label} must contain 1-${max} values`);
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

async function validateStagedSurface(root, sidecar, referenceProgramAssets = []) {
  const rootNames = (await fs.readdir(root)).sort();
  if (JSON.stringify(rootNames) !== JSON.stringify(["SKILL.md", "agents", "artifact-template.json", "assets"].sort())) {
    throw new Error("generated template root does not match the fixed surface");
  }
  const agentNames = await fs.readdir(path.join(root, "agents"));
  const assetNames = (await fs.readdir(path.join(root, "assets"))).sort();
  const exampleNames = (await fs.readdir(path.join(root, "assets", "examples"))).sort();
  if (agentNames.length !== 1 || agentNames[0] !== "agent.yaml") throw new Error("generated agents surface is invalid");
  const expectedAssets = ["examples", "preview.png", ...(sidecar.referenceProgram == null && sidecar.referencePptx == null ? [] : ["references"])].sort();
  if (JSON.stringify(assetNames) !== JSON.stringify(expectedAssets)) throw new Error("generated assets surface is invalid");
  if (JSON.stringify(exampleNames) !== JSON.stringify(sidecar.examples.map((entry) => path.basename(entry.path)).sort())) {
    throw new Error("generated example surface is invalid");
  }
  if (expectedAssets.includes("references")) {
    const referenceNames = (await listRelativeFiles(path.join(root, "assets", "references"))).sort();
    const expectedReferences = [
      ...(sidecar.referenceProgram == null ? [] : ["reference.ppj"]),
      ...(sidecar.referencePptx == null ? [] : ["reference.pptx"]),
      ...referenceProgramAssets.map((asset) => asset.relativePath),
    ].sort();
    if (JSON.stringify(referenceNames) !== JSON.stringify(expectedReferences)) {
      throw new Error("generated reference surface is invalid");
    }
  }
}

async function readReferenceProgramAssets(program, programPath) {
  if (program.assets != null && !Array.isArray(program.assets)) {
    throw new Error("referenceProgram assets must be an array");
  }
  const programRoot = path.dirname(path.resolve(programPath));
  const seen = new Set();
  const assets = [];
  let totalBytes = 0;
  const declared = [];
  if (program.source != null) {
    declared.push({
      index: "source",
      uri: program.source.uri,
      sha256: program.source.sha256,
      label: "source package",
    });
  }
  for (const [index, asset] of (program.assets ?? []).entries()) {
    declared.push({
      index,
      uri: asset?.uri,
      sha256: asset?.sha256,
      label: `assets[${index}]`,
    });
  }
  for (const entry of declared) {
    const { index, uri, sha256: expectedHash, label } = entry;
    if (typeof uri !== "string" || uri.length === 0 || uri.includes("\\") || path.posix.isAbsolute(uri)) {
      throw new Error(`referenceProgram ${label}.uri must be a safe relative path`);
    }
    const segments = uri.split("/");
    if (segments.some((segment) => segment.length === 0 || segment === "." || segment === "..")) {
      throw new Error(`referenceProgram ${label}.uri must not contain empty or traversal segments`);
    }
    if (seen.has(uri)) throw new Error(`referenceProgram asset URI is duplicated: ${uri}`);
    seen.add(uri);
    if (!HASH.test(expectedHash ?? "")) {
      throw new Error(`referenceProgram ${label}.sha256 must be a lowercase SHA-256 value`);
    }
    const maxBytes = label === "source package" ? MAX_REFERENCE_PPTX_BYTES : MAX_REFERENCE_ASSET_BYTES;
    const bytes = await readRegularFile(path.join(programRoot, ...segments), maxBytes, `referenceProgram ${label}`);
    if (sha256(bytes) !== expectedHash) throw new Error(`referenceProgram asset hash does not match: ${uri}`);
    totalBytes += bytes.byteLength;
    if (totalBytes > MAX_TOTAL_REFERENCE_ASSET_BYTES) {
      throw new Error(`referenceProgram assets exceed the ${MAX_TOTAL_REFERENCE_ASSET_BYTES}-byte total budget`);
    }
    assets.push({ relativePath: uri, bytes });
  }
  return assets;
}

async function listRelativeFiles(root, prefix = "") {
  const names = (await fs.readdir(root)).sort();
  const files = [];
  for (const name of names) {
    const absolutePath = path.join(root, name);
    const relativePath = prefix === "" ? name : `${prefix}/${name}`;
    const stat = await fs.lstat(absolutePath);
    if (stat.isSymbolicLink()) throw new Error("generated reference surface must not contain symlinks");
    if (stat.isDirectory()) files.push(...await listRelativeFiles(absolutePath, relativePath));
    else if (stat.isFile()) files.push(relativePath);
    else throw new Error("generated reference surface contains an unsupported entry");
  }
  return files;
}

function validateReferenceSpec(value, label, extension) {
  if (value == null) return;
  if (typeof value !== "object" || Array.isArray(value)) throw new Error(`${label} must be an object`);
  assertOptionalKeys(value, label, ["path", "license", "source", "download"]);
  assertAbsolutePath(value.path, `${label}.path`);
  if (path.extname(value.path).toLowerCase() !== extension) throw new Error(`${label}.path must use ${extension}`);
  assertLine(value.license, `${label}.license`, 120);
  assertLine(value.source, `${label}.source`, 500);
  if (value.download != null) validateRemoteReference(value.download, `${label}.download`);
}

function validateRemoteReference(value, label) {
  if (typeof value !== "object" || Array.isArray(value)) throw new Error(`${label} must be an object`);
  assertKeys(value, label, ["url", "sha256", "bytes"]);
  if (typeof value.url !== "string" || value.url.length > 2048 || /[\0\r\n]/u.test(value.url)) {
    throw new Error(`${label}.url must be a bounded HTTPS URL`);
  }
  let parsed;
  try { parsed = new URL(value.url); } catch { throw new Error(`${label}.url must be a valid HTTPS URL`); }
  if (parsed.protocol !== "https:" || parsed.username || parsed.password || parsed.search || parsed.hash ||
      !REMOTE_REFERENCE_HOSTS.has(parsed.hostname.toLowerCase()) || parsed.pathname.includes("..")) {
    throw new Error(`${label}.url must be an HTTPS raw.githubusercontent.com URL without credentials or traversal`);
  }
  if (!HASH.test(value.sha256 ?? "")) throw new Error(`${label}.sha256 must be a lowercase SHA-256 value`);
  if (!Number.isSafeInteger(value.bytes) || value.bytes < 1 || value.bytes > MAX_REFERENCE_PPTX_BYTES) {
    throw new Error(`${label}.bytes must be a positive bounded integer`);
  }
}

async function readReference(value, { label, maxBytes, extension }) {
  if (value == null) return null;
  if (path.extname(value.path).toLowerCase() !== extension) throw new Error(`${label}.path must use ${extension}`);
  return { ...value, bytes: await readRegularFile(value.path, maxBytes, `${label}.path`) };
}

function referenceMetadata(value, relativePath) {
  return {
    path: relativePath,
    sha256: sha256(value.bytes),
    license: value.license,
    source: value.source,
    ...(value.download == null ? {} : {
      download: {
        ...value.download,
        sha256: sha256(value.bytes),
        bytes: value.bytes.byteLength,
      },
    }),
  };
}

async function inspectCurrentTarget(targetPath) {
  const stat = await fs.lstat(targetPath).catch((error) => {
    if (error?.code === "ENOENT") return null;
    throw error;
  });
  if (stat == null) return null;
  if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error("existing target must be a real directory");
  const sidecarPath = path.join(targetPath, "artifact-template.json");
  const bytes = await readRegularFile(sidecarPath, MAX_SPEC_BYTES, "existing sidecar");
  return { sidecarSha256: sha256(bytes) };
}

async function publishAtomically({ targetPath, stagedPath, current }) {
  if (current == null) {
    await fs.rename(stagedPath, targetPath);
    return;
  }
  const backupPath = `${targetPath}.backup-${randomUUID()}`;
  await fs.rename(targetPath, backupPath);
  try {
    await fs.rename(stagedPath, targetPath);
  } catch (error) {
    await fs.rename(backupPath, targetPath);
    throw error;
  }
  await fs.rm(backupPath, { recursive: true, force: true });
}

async function acquireLock(lockPath) {
  try {
    await fs.mkdir(lockPath, { mode: 0o700 });
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error("another presentation template write is in progress");
    throw error;
  }
}

async function readJsonFile(filePath, maxBytes, label) {
  const bytes = await readRegularFile(filePath, maxBytes, label);
  try {
    return JSON.parse(bytes.toString("utf8"));
  } catch (error) {
    throw new Error(`${label} is not valid JSON: ${error.message}`);
  }
}

async function readTextFile(filePath, maxBytes, label) {
  return (await readRegularFile(filePath, maxBytes, label)).toString("utf8");
}

async function readRegularFile(filePath, maxBytes, label) {
  const absolutePath = path.resolve(filePath);
  const stat = await fs.lstat(absolutePath);
  if (!stat.isFile() || stat.isSymbolicLink()) throw new Error(`${label} must be a regular nonsymlink file`);
  if (stat.size > maxBytes) throw new Error(`${label} exceeds the ${maxBytes}-byte budget`);
  return fs.readFile(absolutePath);
}

async function writeImmutable(filePath, bytes) {
  await fs.writeFile(filePath, bytes, { mode: 0o444, flag: "wx" });
}

function parseArguments(argv) {
  const result = { json: false };
  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index];
    if (flag === "--json") {
      result.json = true;
      continue;
    }
    const value = argv[index + 1];
    if (value == null || value.startsWith("--")) throw new Error(usage());
    index += 1;
    if (flag === "--spec") result.specPath = path.resolve(value);
    else if (flag === "--output-root") result.outputRoot = path.resolve(value);
    else if (flag === "--expected-sha256") result.expectedSha256 = value;
    else throw new Error(usage());
  }
  if (result.specPath == null || result.outputRoot == null) throw new Error(usage());
  if (result.expectedSha256 != null && !HASH.test(result.expectedSha256)) {
    throw new Error("--expected-sha256 must be a lowercase SHA-256 value");
  }
  return result;
}

function usage() {
  return "Usage: package-presentation-template.mjs --spec <absolute.json> --output-root <absolute-dir> [--expected-sha256 <sha256>] [--json]";
}

function assertKeys(value, label, allowedKeys) {
  assertKeysWithOptional(value, label, allowedKeys, []);
}

function assertKeysWithOptional(value, label, requiredKeys, optionalKeys) {
  const allowed = new Set([...requiredKeys, ...optionalKeys]);
  const extras = Object.keys(value).filter((key) => !allowed.has(key));
  if (extras.length > 0) throw new Error(`${label} contains unsupported fields: ${extras.join(", ")}`);
  for (const key of requiredKeys) {
    if (!(key in value)) throw new Error(`${label} is missing ${key}`);
  }
}

function assertOptionalKeys(value, label, allowedKeys) {
  const allowed = new Set(allowedKeys);
  const extras = Object.keys(value).filter((key) => !allowed.has(key));
  if (extras.length > 0) throw new Error(`${label} contains unsupported fields: ${extras.join(", ")}`);
}

function assertLine(value, label, max) {
  if (typeof value !== "string" || value.trim() !== value || value.length === 0 || value.length > max || /[\0\r\n]/u.test(value)) {
    throw new Error(`${label} must be one trimmed line of at most ${max} characters`);
  }
}

function assertEnglishArray(value, label, min, max) {
  if (!Array.isArray(value) || value.length < min || value.length > max) {
    throw new Error(`${label} must contain ${min}-${max} strings`);
  }
  const seen = new Set();
  for (const entry of value) {
    assertLine(entry, label, 120);
    if (!/^[\x20-\x7e]+$/u.test(entry) || !/[a-z]/iu.test(entry)) throw new Error(`${label} must use English search text`);
    const normalized = entry.toLowerCase();
    if (seen.has(normalized)) throw new Error(`${label} must not contain duplicates`);
    seen.add(normalized);
  }
}

function assertAbsolutePath(value, label) {
  if (typeof value !== "string" || !path.isAbsolute(value) || /[\0\r\n]/u.test(value)) {
    throw new Error(`${label} must be an absolute path`);
  }
}

function assertEnum(value, label, allowed) {
  if (!allowed.has(value)) throw new Error(`${label} must be one of ${[...allowed].join(", ")}`);
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function createMontage(images) {
  const columns = 2;
  const rows = Math.ceil(images.length / columns);
  const margin = 28;
  const gap = 20;
  const tileWidth = 602;
  const tileHeight = 339;
  const width = margin * 2 + columns * tileWidth + gap;
  const height = margin * 2 + rows * tileHeight + (rows - 1) * gap;
  const pixels = new Uint8Array(width * height * 4);
  for (let offset = 0; offset < pixels.length; offset += 4) {
    pixels[offset] = 239;
    pixels[offset + 1] = 241;
    pixels[offset + 2] = 245;
    pixels[offset + 3] = 255;
  }
  images.forEach((image, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const x = margin + column * (tileWidth + gap);
    const y = margin + row * (tileHeight + gap);
    blitContained(image, { width, height, pixels }, x, y, tileWidth, tileHeight);
  });
  return { width, height, pixels };
}

function blitContained(source, destination, x, y, boxWidth, boxHeight) {
  const scale = Math.min(boxWidth / source.width, boxHeight / source.height);
  const drawWidth = Math.max(1, Math.round(source.width * scale));
  const drawHeight = Math.max(1, Math.round(source.height * scale));
  const offsetX = x + Math.floor((boxWidth - drawWidth) / 2);
  const offsetY = y + Math.floor((boxHeight - drawHeight) / 2);
  for (let py = 0; py < drawHeight; py += 1) {
    const sourceY = Math.min(source.height - 1, Math.floor(py / scale));
    for (let px = 0; px < drawWidth; px += 1) {
      const sourceX = Math.min(source.width - 1, Math.floor(px / scale));
      const sourceOffset = (sourceY * source.width + sourceX) * 4;
      const destinationOffset = ((offsetY + py) * destination.width + offsetX + px) * 4;
      destination.pixels.set(source.pixels.subarray(sourceOffset, sourceOffset + 4), destinationOffset);
    }
  }
}

function decodePngRgba(bytes) {
  if (!Buffer.from(bytes.subarray(0, 8)).equals(PNG_SIGNATURE)) throw new Error("example is not a PNG file");
  let offset = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let compression = 0;
  let filterMethod = 0;
  let interlace = 0;
  let palette = null;
  let paletteAlpha = null;
  const idat = [];
  while (offset + 12 <= bytes.byteLength) {
    const length = bytes.readUInt32BE(offset);
    const type = bytes.toString("ascii", offset + 4, offset + 8);
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    if (dataEnd + 4 > bytes.byteLength) throw new Error("example PNG chunk is truncated");
    const data = bytes.subarray(dataStart, dataEnd);
    if (type === "IHDR") {
      width = data.readUInt32BE(0);
      height = data.readUInt32BE(4);
      bitDepth = data[8];
      colorType = data[9];
      compression = data[10];
      filterMethod = data[11];
      interlace = data[12];
    } else if (type === "PLTE") palette = Buffer.from(data);
    else if (type === "tRNS") paletteAlpha = Buffer.from(data);
    else if (type === "IDAT") idat.push(data);
    else if (type === "IEND") break;
    offset = dataEnd + 4;
  }
  if (!width || !height) throw new Error("example PNG is missing IHDR geometry");
  if (bitDepth !== 8 || compression !== 0 || filterMethod !== 0 || interlace !== 0) {
    throw new Error("examples must be 8-bit non-interlaced PNGs");
  }
  const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[colorType];
  if (!channels) throw new Error(`unsupported example PNG color type ${colorType}`);
  if (colorType === 3 && (palette == null || palette.length === 0 || palette.length % 3 !== 0)) {
    throw new Error("indexed example PNG is missing a valid palette");
  }
  const rowBytes = width * channels;
  const inflated = inflateSync(Buffer.concat(idat));
  if (inflated.byteLength < (rowBytes + 1) * height) throw new Error("example PNG image data is truncated");
  const raw = new Uint8Array(width * height * channels);
  let inputOffset = 0;
  for (let py = 0; py < height; py += 1) {
    const filter = inflated[inputOffset++];
    const rowStart = py * rowBytes;
    const priorStart = (py - 1) * rowBytes;
    for (let px = 0; px < rowBytes; px += 1) {
      const left = px >= channels ? raw[rowStart + px - channels] : 0;
      const up = py > 0 ? raw[priorStart + px] : 0;
      const upLeft = py > 0 && px >= channels ? raw[priorStart + px - channels] : 0;
      const value = inflated[inputOffset++];
      if (filter === 0) raw[rowStart + px] = value;
      else if (filter === 1) raw[rowStart + px] = (value + left) & 0xff;
      else if (filter === 2) raw[rowStart + px] = (value + up) & 0xff;
      else if (filter === 3) raw[rowStart + px] = (value + Math.floor((left + up) / 2)) & 0xff;
      else if (filter === 4) raw[rowStart + px] = (value + paeth(left, up, upLeft)) & 0xff;
      else throw new Error(`unsupported example PNG row filter ${filter}`);
    }
  }
  const pixels = new Uint8Array(width * height * 4);
  for (let sourceOffset = 0, pixel = 0; pixel < width * height; pixel += 1, sourceOffset += channels) {
    const destinationOffset = pixel * 4;
    if (colorType === 0) {
      pixels[destinationOffset] = raw[sourceOffset];
      pixels[destinationOffset + 1] = raw[sourceOffset];
      pixels[destinationOffset + 2] = raw[sourceOffset];
      pixels[destinationOffset + 3] = 255;
    } else if (colorType === 2) {
      pixels.set([raw[sourceOffset], raw[sourceOffset + 1], raw[sourceOffset + 2], 255], destinationOffset);
    } else if (colorType === 3) {
      const paletteIndex = raw[sourceOffset];
      const paletteOffset = paletteIndex * 3;
      if (paletteOffset + 2 >= palette.length) throw new Error("indexed example PNG references a missing palette entry");
      pixels.set([
        palette[paletteOffset],
        palette[paletteOffset + 1],
        palette[paletteOffset + 2],
        paletteAlpha?.[paletteIndex] ?? 255,
      ], destinationOffset);
    } else if (colorType === 4) {
      pixels.set([raw[sourceOffset], raw[sourceOffset], raw[sourceOffset], raw[sourceOffset + 1]], destinationOffset);
    } else pixels.set(raw.subarray(sourceOffset, sourceOffset + 4), destinationOffset);
  }
  return { width, height, pixels };
}

function paeth(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  return pb <= pc ? b : c;
}

function encodePngRgba({ width, height, pixels }) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const rowBytes = width * 4;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let row = 0; row < height; row += 1) {
    Buffer.from(pixels.subarray(row * rowBytes, (row + 1) * rowBytes))
      .copy(raw, row * (rowBytes + 1) + 1);
  }
  return Buffer.concat([
    PNG_SIGNATURE,
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", deflateSync(raw)),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

function pngChunk(type, data) {
  const chunk = Buffer.alloc(12 + data.length);
  chunk.writeUInt32BE(data.length, 0);
  chunk.write(type, 4, 4, "ascii");
  data.copy(chunk, 8);
  chunk.writeUInt32BE(crc32(chunk.subarray(4, 8 + data.length)), 8 + data.length);
  return chunk;
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
