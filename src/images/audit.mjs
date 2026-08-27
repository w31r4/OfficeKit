import { createHash, randomUUID } from "node:crypto";
import { link, lstat, open, readFile, realpath, rm } from "node:fs/promises";
import path from "node:path";

import { loadOoxmlZipWithinBudget } from "../ooxml/package.mjs";
import { imageContentTypeFromExtension } from "../shared/images.mjs";
import { imageError } from "./errors.mjs";
import { listTaskImageAssets, writeTaskImageAuditEvidence } from "./task-assets.mjs";

export const IMAGE_AUDIT_SCHEMA = "office-kit/presentation-image-audit/v1";
const MAX_PPTX_BYTES = 256 * 1024 * 1024;
const IMAGE_EXTENSIONS = new Set(["png", "jpg", "jpeg", "gif", "svg", "webp", "bmp", "tif", "tiff", "emf", "wmf"]);

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function regularFile(target, label, maximum = MAX_PPTX_BYTES) {
  let stat;
  try { stat = await lstat(target); }
  catch (error) {
    if (error?.code === "ENOENT") throw imageError("image-audit-input-missing", `${label} does not exist: ${target}`);
    throw error;
  }
  if (stat.isSymbolicLink() || !stat.isFile()) throw imageError("unsafe-image-path", `${label} must be a regular non-symlink file.`);
  if (stat.size > maximum) throw imageError("image-audit-input-too-large", `${label} exceeds ${maximum} bytes.`);
  return stat;
}

function mimeForPart(partPath) {
  const extension = path.posix.extname(partPath).slice(1).toLowerCase();
  const canonical = imageContentTypeFromExtension(extension);
  if (canonical !== "application/octet-stream") return canonical;
  return extension ? `image/${extension}` : "application/octet-stream";
}

function assetAuditRecord(asset, parts) {
  return {
    sha256: asset.sha256,
    parts: [...parts].sort(),
    mimeType: asset.mimeType,
    width: asset.width,
    height: asset.height,
    rights: asset.rights,
    provenance: asset.provenance,
    source: asset.source,
    creditLine: asset.creditLine,
    visibleAttributionRequired: asset.visibleAttributionRequired,
  };
}

export async function auditPresentationImages(task, input = {}) {
  const pptxPath = path.resolve(String(input.pptxPath || ""));
  await regularFile(pptxPath, "Presentation image audit input");
  const canonicalPptxPath = await realpath(pptxPath);
  const pptxBytes = await readFile(canonicalPptxPath);
  const presentationSha256 = sha256(pptxBytes);
  const zip = await loadOoxmlZipWithinBudget(pptxBytes, {
    maxInputBytes: MAX_PPTX_BYTES,
    maxParts: 5_000,
    maxPartBytes: 64 * 1024 * 1024,
    maxTotalBytes: 256 * 1024 * 1024,
    maxCompressionRatio: 250,
  }, "PPTX image audit");
  const media = [];
  for (const entry of Object.values(zip.files).filter((file) => !file.dir && /^ppt\/media\/[^/]+$/u.test(file.name)).sort((left, right) => left.name.localeCompare(right.name))) {
    const extension = path.posix.extname(entry.name).slice(1).toLowerCase();
    if (!IMAGE_EXTENSIONS.has(extension)) continue;
    const bytes = Buffer.from(await entry.async("uint8array"));
    media.push({ partPath: entry.name, sha256: sha256(bytes), byteLength: bytes.length, mimeType: mimeForPart(entry.name) });
  }

  const assets = await listTaskImageAssets(task);
  const mediaByHash = new Map();
  for (const item of media) {
    const parts = mediaByHash.get(item.sha256) || [];
    parts.push(item.partPath);
    mediaByHash.set(item.sha256, parts);
  }
  const assetByHash = new Map(assets.map((asset) => [asset.sha256, asset]));
  const used = assets
    .filter((asset) => mediaByHash.has(asset.sha256))
    .map((asset) => assetAuditRecord(asset, mediaByHash.get(asset.sha256)))
    .sort((left, right) => left.sha256.localeCompare(right.sha256));
  const unused = assets
    .filter((asset) => !mediaByHash.has(asset.sha256))
    .map((asset) => assetAuditRecord(asset, []))
    .sort((left, right) => left.sha256.localeCompare(right.sha256));
  const unregistered = media
    .filter((item) => !assetByHash.has(item.sha256))
    .sort((left, right) => left.partPath.localeCompare(right.partPath));
  const attributions = used
    .filter((asset) => asset.creditLine)
    .map((asset) => ({ sha256: asset.sha256, parts: asset.parts, rights: asset.rights, creditLine: asset.creditLine, visibleRequired: asset.visibleAttributionRequired }))
    .sort((left, right) => left.sha256.localeCompare(right.sha256));
  const report = {
    schema: IMAGE_AUDIT_SCHEMA,
    taskId: task.manifest.id,
    presentation: { path: canonicalPptxPath, sha256: presentationSha256, byteLength: pptxBytes.length, imagePartCount: media.length },
    used,
    unused,
    unregistered,
    attributions,
    ok: unregistered.length === 0 && attributions.every((item) => item.visibleRequired !== true || item.creditLine),
  };
  const evidence = await writeTaskImageAuditEvidence(task, report);
  return Object.freeze({ ...report, evidence });
}

async function canonicalOutputPath(target) {
  const absolute = path.resolve(target);
  const parent = await realpath(path.dirname(absolute));
  const candidate = path.join(parent, path.basename(absolute));
  try {
    const stat = await lstat(candidate);
    if (stat.isSymbolicLink() || !stat.isFile()) throw imageError("unsafe-image-path", "Sources output must be a regular non-symlink file path.");
    return { path: await realpath(candidate), exists: true };
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
    return { path: candidate, exists: false };
  }
}

export async function writeImageSourcesSidecar(report, outputPath) {
  const input = await realpath(report.presentation.path);
  const output = await canonicalOutputPath(outputPath);
  if (output.path === input) throw imageError("unsafe-image-output", "Sources output must not overwrite the presentation input.");
  if (output.exists) throw imageError("image-output-exists", `Sources output already exists: ${output.path}`);
  const sidecar = {
    schema: IMAGE_AUDIT_SCHEMA,
    taskId: report.taskId,
    presentation: { sha256: report.presentation.sha256, byteLength: report.presentation.byteLength, imagePartCount: report.presentation.imagePartCount },
    used: report.used,
    unused: report.unused,
    unregistered: report.unregistered,
    attributions: report.attributions,
    ok: report.ok,
  };
  const bytes = Buffer.from(`${JSON.stringify(sidecar, null, 2)}\n`);
  const temporary = path.join(path.dirname(output.path), `.${path.basename(output.path)}.${randomUUID()}.tmp`);
  const handle = await open(temporary, "wx", 0o600);
  try {
    await handle.writeFile(bytes);
    await handle.sync();
  } finally {
    await handle.close();
  }
  try {
    await link(temporary, output.path);
  } catch (error) {
    if (error?.code === "EEXIST") {
      throw imageError("image-output-exists", `Sources output already exists: ${output.path}`);
    }
    throw error;
  } finally {
    await rm(temporary, { force: true });
  }
  return Object.freeze({ path: output.path, sha256: sha256(bytes), byteLength: bytes.length });
}
