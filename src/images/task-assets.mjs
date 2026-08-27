import { createHash, randomUUID } from "node:crypto";
import {
  chmod,
  link,
  lstat,
  mkdir,
  open,
  readFile,
  readdir,
  realpath,
  rm,
} from "node:fs/promises";
import path from "node:path";

import { openTask } from "../cli/task-store.mjs";
import { inspectImageBytes, normalizeImageMimeType } from "../shared/image-bytes.mjs";
import { imageError } from "./errors.mjs";
import { imageRightsCompatible, normalizeImageRights } from "./rights.mjs";

export const IMAGE_ASSET_RECEIPT_SCHEMA = "office-kit/image-asset-receipt/v1";
export const IMAGE_SEARCH_EVIDENCE_SCHEMA = "office-kit/image-search-evidence/v1";
export const MAX_IMAGE_BYTES = 20 * 1024 * 1024;
export const MAX_IMAGE_PIXELS = 40_000_000;
export const MAX_IMAGE_DIMENSION = 16_384;

const SHA256_PATTERN = /^[0-9a-f]{64}$/u;
const SEARCH_ID_PATTERN = /^s_[0-9a-f]{16}$/u;
const CANDIDATE_REF_PATTERN = /^imgc_[0-9a-f]{24}$/u;
const MAX_RECEIPT_BYTES = 128 * 1024;
const MAX_SEARCH_BYTES = 2 * 1024 * 1024;
const MAX_SEARCHES = 500;

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function contained(candidate, root, label) {
  const relative = path.relative(root, candidate);
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw imageError("unsafe-image-path", `${label} escapes its task root.`);
  }
}

async function lstatIfExists(target) {
  try { return await lstat(target); }
  catch (error) {
    if (error?.code === "ENOENT") return null;
    throw error;
  }
}

async function ensurePrivateDirectory(target, root) {
  contained(path.resolve(target), path.resolve(root), "Image directory");
  const existing = await lstatIfExists(target);
  if (existing?.isSymbolicLink() || (existing && !existing.isDirectory())) {
    throw imageError("unsafe-image-path", `Image path must be a regular directory: ${target}`);
  }
  if (!existing) await mkdir(target, { mode: 0o700 });
  const canonical = await realpath(target);
  const canonicalRoot = await realpath(root);
  contained(canonical, canonicalRoot, "Image directory");
  await chmod(canonical, 0o700);
  return canonical;
}

async function ensureImageDirectories(task) {
  const assets = await ensurePrivateDirectory(path.join(task.taskRoot, "assets"), task.taskRoot);
  const imageAssets = await ensurePrivateDirectory(path.join(assets, "images"), task.taskRoot);
  const evidenceImages = await ensurePrivateDirectory(path.join(task.taskRoot, "evidence", "images"), task.taskRoot);
  const searches = await ensurePrivateDirectory(path.join(evidenceImages, "searches"), task.taskRoot);
  const audits = await ensurePrivateDirectory(path.join(evidenceImages, "audits"), task.taskRoot);
  return { imageAssets, evidenceImages, searches, audits };
}

async function publishImmutable(target, bytes, mode) {
  const existing = await lstatIfExists(target);
  if (existing) {
    if (existing.isSymbolicLink() || !existing.isFile()) throw imageError("unsafe-image-path", `Image file must be a regular non-symlink file: ${target}`);
    const current = await readFile(target);
    if (sha256(current) !== sha256(bytes)) throw imageError("image-asset-collision", `Immutable image path contains different bytes: ${target}`);
    await chmod(target, mode);
    return false;
  }

  const temporary = path.join(path.dirname(target), `.${path.basename(target)}.${randomUUID()}.tmp`);
  const handle = await open(temporary, "wx", mode);
  try {
    await handle.writeFile(bytes);
    await handle.sync();
  } finally {
    await handle.close();
  }
  try {
    await link(temporary, target);
  } catch (error) {
    if (error?.code !== "EEXIST") throw error;
    const current = await readFile(target);
    if (sha256(current) !== sha256(bytes)) throw imageError("image-asset-collision", `Immutable image path contains different bytes: ${target}`);
  } finally {
    await rm(temporary, { force: true });
  }
  await chmod(target, mode);
  return true;
}

async function readBoundedJson(target, maximum, label) {
  const stat = await lstatIfExists(target);
  if (!stat || stat.isSymbolicLink() || !stat.isFile()) throw imageError("unsafe-image-path", `${label} must be a regular non-symlink file.`);
  if (stat.size > maximum) throw imageError("image-evidence-too-large", `${label} exceeds ${maximum} bytes.`);
  try { return JSON.parse(await readFile(target, "utf8")); }
  catch { throw imageError("invalid-image-evidence", `${label} is not valid JSON.`); }
}

function receiptDescriptor(receipt, taskRoot) {
  return Object.freeze({
    schema: receipt.schema,
    sha256: receipt.sha256,
    path: path.resolve(taskRoot, receipt.assetPath),
    receiptPath: path.resolve(taskRoot, receipt.receiptPath),
    mimeType: receipt.mimeType,
    extension: receipt.extension,
    width: receipt.width,
    height: receipt.height,
    byteLength: receipt.byteLength,
    rights: receipt.rights.rights,
    provenance: receipt.rights,
    source: receipt.source,
    creditLine: receipt.rights.creditLine,
    visibleAttributionRequired: receipt.rights.visibleAttributionRequired,
    createdAt: receipt.createdAt,
  });
}

async function validateReceipt(receipt, taskRoot) {
  if (receipt?.schema !== IMAGE_ASSET_RECEIPT_SCHEMA || !SHA256_PATTERN.test(String(receipt?.sha256 || ""))) {
    throw imageError("invalid-image-receipt", "Image receipt schema or SHA-256 is invalid.");
  }
  const expectedAsset = `assets/images/${receipt.sha256}.${receipt.extension}`;
  const expectedReceipt = `assets/images/${receipt.sha256}.json`;
  if (receipt.assetPath !== expectedAsset || receipt.receiptPath !== expectedReceipt) {
    throw imageError("invalid-image-receipt", "Image receipt paths are not canonical.");
  }
  const assetPath = path.resolve(taskRoot, receipt.assetPath);
  contained(assetPath, taskRoot, "Image asset");
  const stat = await lstatIfExists(assetPath);
  if (!stat || stat.isSymbolicLink() || !stat.isFile() || stat.size !== receipt.byteLength) {
    throw imageError("image-asset-corrupt", `Image asset ${receipt.sha256} is missing or unsafe.`);
  }
  const bytes = await readFile(assetPath);
  if (sha256(bytes) !== receipt.sha256) throw imageError("image-asset-corrupt", `Image asset ${receipt.sha256} does not match its receipt.`);
  const inspected = inspectImageBytes(bytes, {
    declaredMimeType: receipt.mimeType,
    label: `Task image ${receipt.sha256}`,
    maxBytes: MAX_IMAGE_BYTES,
    maxPixels: MAX_IMAGE_PIXELS,
    maxDimension: MAX_IMAGE_DIMENSION,
  });
  if (inspected.extension !== receipt.extension || inspected.width !== receipt.width || inspected.height !== receipt.height) {
    throw imageError("image-asset-corrupt", `Image asset ${receipt.sha256} dimensions or type do not match its receipt.`);
  }
  return receiptDescriptor(receipt, taskRoot);
}

export async function openImageTask({ workspaceRoot, taskId }) {
  return openTask({ workspaceRoot, taskId });
}

export async function addTaskImageAsset(task, input = {}) {
  const bytes = Buffer.from(input.bytes || []);
  const declaredMimeType = normalizeImageMimeType(input.mimeType);
  const inspected = inspectImageBytes(bytes, {
    declaredMimeType,
    label: "Task image asset",
    maxBytes: MAX_IMAGE_BYTES,
    maxPixels: MAX_IMAGE_PIXELS,
    maxDimension: MAX_IMAGE_DIMENSION,
  });
  const digest = sha256(bytes);
  const rights = normalizeImageRights(input.rights, input.rightsMetadata);
  const source = input.source && typeof input.source === "object" && !Array.isArray(input.source)
    ? structuredClone(input.source)
    : { kind: "unknown" };
  const directories = await ensureImageDirectories(task);
  const assetPath = path.join(directories.imageAssets, `${digest}.${inspected.extension}`);
  const receiptPath = path.join(directories.imageAssets, `${digest}.json`);
  const existingReceipt = await lstatIfExists(receiptPath);
  if (existingReceipt) {
    const receipt = await readBoundedJson(receiptPath, MAX_RECEIPT_BYTES, "Image receipt");
    const descriptor = await validateReceipt(receipt, task.taskRoot);
    if (!imageRightsCompatible(receipt.rights, rights) || JSON.stringify(receipt.source) !== JSON.stringify(source)) {
      throw imageError("image-provenance-conflict", `Image ${digest} is already registered with different provenance.`);
    }
    return descriptor;
  }

  await publishImmutable(assetPath, bytes, 0o400);
  const receipt = {
    schema: IMAGE_ASSET_RECEIPT_SCHEMA,
    sha256: digest,
    assetPath: path.relative(task.taskRoot, assetPath).split(path.sep).join("/"),
    receiptPath: path.relative(task.taskRoot, receiptPath).split(path.sep).join("/"),
    mimeType: inspected.mimeType,
    extension: inspected.extension,
    width: inspected.width,
    height: inspected.height,
    byteLength: inspected.byteLength,
    rights,
    source,
    createdAt: (input.now || new Date()).toISOString(),
  };
  const receiptBytes = Buffer.from(`${JSON.stringify(receipt, null, 2)}\n`);
  await publishImmutable(receiptPath, receiptBytes, 0o600);
  return validateReceipt(receipt, task.taskRoot);
}

export async function listTaskImageAssets(task) {
  const root = path.join(task.taskRoot, "assets", "images");
  const stat = await lstatIfExists(root);
  if (!stat) return [];
  if (stat.isSymbolicLink() || !stat.isDirectory()) throw imageError("unsafe-image-path", "Task image assets path is unsafe.");
  const entries = await readdir(root, { withFileTypes: true });
  const receipts = [];
  for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
    if (!entry.name.endsWith(".json")) continue;
    if (!entry.isFile() || entry.isSymbolicLink() || !SHA256_PATTERN.test(entry.name.slice(0, -5))) {
      throw imageError("invalid-image-receipt", `Unexpected image receipt entry: ${entry.name}`);
    }
    const receipt = await readBoundedJson(path.join(root, entry.name), MAX_RECEIPT_BYTES, "Image receipt");
    receipts.push(await validateReceipt(receipt, task.taskRoot));
  }
  return receipts;
}

function publicCandidate(candidate) {
  return Object.freeze({
    candidateRef: candidate.candidateRef,
    provider: candidate.provider,
    kind: candidate.kind,
    title: candidate.title,
    author: candidate.author,
    previewUrl: candidate.previewUrl,
    sourcePage: candidate.sourcePage,
    width: candidate.width,
    height: candidate.height,
    mimeType: candidate.mimeType,
    rights: candidate.rights?.rights,
    rightsEvidence: candidate.rights?.evidence,
    creditLine: candidate.rights?.creditLine,
    visibleAttributionRequired: candidate.rights?.visibleAttributionRequired === true,
    score: candidate.score,
  });
}

export async function recordTaskImageSearch(task, input = {}) {
  const directories = await ensureImageDirectories(task);
  const searchId = `s_${randomUUID().replaceAll("-", "").slice(0, 16)}`;
  const searchRoot = await ensurePrivateDirectory(path.join(directories.searches, searchId), task.taskRoot);
  const candidates = (input.candidates || []).map((candidate) => ({
    ...structuredClone(candidate),
    candidateRef: `imgc_${randomUUID().replaceAll("-", "").slice(0, 24)}`,
  }));
  const record = {
    schema: IMAGE_SEARCH_EVIDENCE_SCHEMA,
    searchId,
    taskId: task.manifest.id,
    query: String(input.query || ""),
    kind: input.kind,
    purpose: input.purpose,
    orientation: input.orientation,
    selectionMade: false,
    candidates,
    rejected: input.rejected || [],
    providerReports: input.providerReports || [],
    createdAt: (input.now || new Date()).toISOString(),
  };
  const evidencePath = path.join(searchRoot, "search.json");
  await publishImmutable(evidencePath, Buffer.from(`${JSON.stringify(record, null, 2)}\n`), 0o600);
  return Object.freeze({
    searchId,
    taskId: task.manifest.id,
    query: record.query,
    kind: record.kind,
    purpose: record.purpose,
    orientation: record.orientation,
    selectionMade: false,
    candidates: candidates.map(publicCandidate),
    rejected: record.rejected,
    providerReports: record.providerReports,
    evidencePath,
  });
}

async function readSearchRecords(task) {
  const root = path.join(task.taskRoot, "evidence", "images", "searches");
  const stat = await lstatIfExists(root);
  if (!stat) return [];
  if (stat.isSymbolicLink() || !stat.isDirectory()) throw imageError("unsafe-image-path", "Task image search path is unsafe.");
  const entries = await readdir(root, { withFileTypes: true });
  if (entries.length > MAX_SEARCHES) throw imageError("image-search-budget-exceeded", `Task contains more than ${MAX_SEARCHES} image searches.`);
  const records = [];
  for (const entry of entries.sort((left, right) => right.name.localeCompare(left.name))) {
    if (!SEARCH_ID_PATTERN.test(entry.name) || !entry.isDirectory() || entry.isSymbolicLink()) {
      throw imageError("invalid-image-evidence", `Unexpected image search entry: ${entry.name}`);
    }
    const target = path.join(root, entry.name, "search.json");
    const record = await readBoundedJson(target, MAX_SEARCH_BYTES, "Image search evidence");
    if (record?.schema !== IMAGE_SEARCH_EVIDENCE_SCHEMA || record.taskId !== task.manifest.id || record.searchId !== entry.name || record.selectionMade !== false || !Array.isArray(record.candidates)) {
      throw imageError("invalid-image-evidence", `Image search ${entry.name} is invalid.`);
    }
    records.push({ record, path: target });
  }
  return records;
}

export async function listTaskImageSearches(task) {
  return (await readSearchRecords(task)).map(({ record, path: evidencePath }) => ({
    searchId: record.searchId,
    query: record.query,
    kind: record.kind,
    purpose: record.purpose,
    orientation: record.orientation,
    selectionMade: false,
    candidates: record.candidates.map(publicCandidate),
    rejected: record.rejected,
    providerReports: record.providerReports,
    createdAt: record.createdAt,
    evidencePath,
  }));
}

export async function resolveTaskImageCandidate(task, candidateRef) {
  const ref = String(candidateRef || "");
  if (!CANDIDATE_REF_PATTERN.test(ref)) throw imageError("invalid-image-candidate", "Image candidate ref is invalid.");
  for (const { record } of await readSearchRecords(task)) {
    const candidate = record.candidates.find((item) => item.candidateRef === ref);
    if (candidate) return structuredClone(candidate);
  }
  throw imageError("image-candidate-not-found", `Image candidate ${ref} does not belong to task ${task.manifest.id}.`);
}

export async function imageTaskState({ workspaceRoot, taskId }) {
  const task = await openImageTask({ workspaceRoot, taskId });
  const [assets, searches] = await Promise.all([listTaskImageAssets(task), listTaskImageSearches(task)]);
  return { task, assets, searches };
}
