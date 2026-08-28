import path from "node:path";

const DEFAULT_MAX_INPUT_BYTES = 256 * 1024 * 1024;
const DEFAULT_MAX_PARTS = 5_000;
const DEFAULT_MAX_PART_BYTES = 64 * 1024 * 1024;
const DEFAULT_MAX_TOTAL_BYTES = 256 * 1024 * 1024;
const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const LOCAL_SIGNATURE = 0x04034b50;
const UTF8_FLAG = 0x0800;
const ENCRYPTED_FLAG = 0x0001;
const ZIP64_U16 = 0xffff;
const ZIP64_U32 = 0xffff_ffff;
const decoder = new TextDecoder();

function positiveLimit(value, fallback) {
  const number = Number(value ?? fallback);
  return Number.isSafeInteger(number) && number > 0 ? number : fallback;
}

function zipLimits(options = {}) {
  const maxCompressionRatio = Number(options.maxCompressionRatio ?? 0);
  return {
    maxInputBytes: positiveLimit(options.maxInputBytes, DEFAULT_MAX_INPUT_BYTES),
    maxParts: positiveLimit(options.maxParts, DEFAULT_MAX_PARTS),
    maxPartBytes: positiveLimit(options.maxPartBytes, DEFAULT_MAX_PART_BYTES),
    maxTotalBytes: positiveLimit(options.maxTotalBytes, DEFAULT_MAX_TOTAL_BYTES),
    maxCompressionRatio: Number.isFinite(maxCompressionRatio) && maxCompressionRatio > 0 ? maxCompressionRatio : 0,
  };
}

function safePartPath(partPath, family) {
  const raw = String(partPath || "").replaceAll("\\", "/").trim();
  if (!raw || raw.startsWith("/") || raw.includes("\0")) throw new Error(`Unsafe ${family} part path: ${partPath}`);
  const normalized = path.posix.normalize(raw).replace(/^\.\//, "");
  if (!normalized || normalized === "." || normalized.startsWith("../") || normalized.includes("/../") || normalized === "..") {
    throw new Error(`Unsafe ${family} part path: ${partPath}`);
  }
  if (normalized.length > 1024) throw new Error(`Unsafe ${family} part path: path exceeds 1024 characters`);
  return normalized;
}

function findEndOfCentralDirectory(view) {
  const minimum = Math.max(0, view.byteLength - 22 - 0xffff);
  for (let offset = view.byteLength - 22; offset >= minimum; offset -= 1) {
    if (view.getUint32(offset, true) !== EOCD_SIGNATURE) continue;
    const commentLength = view.getUint16(offset + 20, true);
    if (offset + 22 + commentLength === view.byteLength) return offset;
  }
  return -1;
}

function asciiOrUtf8(bytes, utf8) {
  if (!utf8 && bytes.some((value) => value >= 0x80)) return undefined;
  return decoder.decode(bytes);
}

// Fast path for ordinary single-disk Office ZIPs. Returning undefined is an
// intentional compatibility signal: callers must use the existing JSZip path
// for ZIP64, prefixed archives, encrypted entries, or legacy path encodings.
export function tryIndexOoxmlZipWithinBudget(value, options = {}, family = "OOXML") {
  const bytes = value instanceof Uint8Array
    ? value
    : ArrayBuffer.isView(value)
      ? new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
      : new Uint8Array(value);
  const limits = zipLimits(options);
  if (bytes.byteLength > limits.maxInputBytes) {
    throw new Error(`${family} input has ${bytes.byteLength} bytes and exceeds maxInputBytes (${limits.maxInputBytes}).`);
  }
  if (bytes.byteLength < 22) throw new Error(`${family} package is not a readable ZIP package.`);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = findEndOfCentralDirectory(view);
  if (eocd < 0) throw new Error(`${family} package is not a readable ZIP package.`);
  const disk = view.getUint16(eocd + 4, true);
  const centralDisk = view.getUint16(eocd + 6, true);
  const entriesOnDisk = view.getUint16(eocd + 8, true);
  const entryCount = view.getUint16(eocd + 10, true);
  const centralSize = view.getUint32(eocd + 12, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  if (disk !== 0 || centralDisk !== 0 || entriesOnDisk !== entryCount || entryCount === ZIP64_U16 ||
      centralSize === ZIP64_U32 || centralOffset === ZIP64_U32 || centralOffset + centralSize !== eocd) {
    return undefined;
  }

  const entries = new Map();
  let cursor = centralOffset;
  let fileCount = 0;
  let declaredTotalBytes = 0;
  for (let index = 0; index < entryCount; index += 1) {
    if (cursor + 46 > eocd || view.getUint32(cursor, true) !== CENTRAL_SIGNATURE) {
      throw new Error(`${family} package has an invalid central directory.`);
    }
    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const uncompressedSize = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const localOffset = view.getUint32(cursor + 42, true);
    const recordEnd = cursor + 46 + nameLength + extraLength + commentLength;
    if (recordEnd > eocd) throw new Error(`${family} package has a truncated central-directory entry.`);
    if ((flags & ENCRYPTED_FLAG) || !new Set([0, 8]).has(method) ||
        compressedSize === ZIP64_U32 || uncompressedSize === ZIP64_U32 || localOffset === ZIP64_U32) {
      return undefined;
    }
    const name = asciiOrUtf8(bytes.subarray(cursor + 46, cursor + 46 + nameLength), Boolean(flags & UTF8_FLAG));
    if (name === undefined) return undefined;
    cursor = recordEnd;
    if (name.endsWith("/")) continue;
    const partPath = safePartPath(name, family);
    if (entries.has(partPath)) return undefined;
    fileCount += 1;
    if (fileCount > limits.maxParts) throw new Error(`${family} package has ${fileCount} parts; maxParts is ${limits.maxParts}.`);
    if (uncompressedSize > limits.maxPartBytes) {
      throw new Error(`${family} part ${partPath} exceeds maxPartBytes (${limits.maxPartBytes}).`);
    }
    declaredTotalBytes += uncompressedSize;
    if (!Number.isSafeInteger(declaredTotalBytes) || declaredTotalBytes > limits.maxTotalBytes) {
      throw new Error(`${family} package exceeds maxTotalBytes (${limits.maxTotalBytes}).`);
    }
    if (limits.maxCompressionRatio > 0 && uncompressedSize > 0) {
      const ratio = compressedSize > 0 ? uncompressedSize / compressedSize : Infinity;
      if (!Number.isFinite(ratio) || ratio > limits.maxCompressionRatio) {
        throw new Error(`${family} part ${partPath} exceeds maxCompressionRatio (${limits.maxCompressionRatio}).`);
      }
    }
    if (localOffset + 30 > centralOffset || view.getUint32(localOffset, true) !== LOCAL_SIGNATURE) {
      throw new Error(`${family} part ${partPath} has an invalid local ZIP header.`);
    }
    const localFlags = view.getUint16(localOffset + 6, true);
    const localMethod = view.getUint16(localOffset + 8, true);
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength;
    if (localFlags !== flags || localMethod !== method || dataOffset + compressedSize > centralOffset) {
      throw new Error(`${family} part ${partPath} has inconsistent local ZIP metadata.`);
    }
    entries.set(partPath, {
      compressedContent: bytes.subarray(dataOffset, dataOffset + compressedSize),
      compressedSize,
      uncompressedSize,
      compression: method === 0 ? "\x00\x00" : "\x08\x00",
    });
  }
  if (cursor !== eocd) throw new Error(`${family} package has trailing central-directory data.`);
  return entries;
}
