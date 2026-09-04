import { decodePngRgba, isPngBytes } from "./png.mjs";

/**
 * A small, bounded visual contract for selecting and reviewing presentation
 * imagery. These values describe what the bytes or a declared inspection
 * prove; they are not a promise that an image is legally cleared or that a
 * subject was semantically recognized.
 */
export const IMAGE_EDGE_QUALITIES = Object.freeze(["clean", "soft", "fringe", "unknown"]);
export const IMAGE_SHADOW_MODES = Object.freeze(["none", "baked", "separate", "unknown"]);

const DEFAULT_PROFILE = Object.freeze({
  alphaPresent: null,
  subjectBounds: null,
  edgeQuality: "unknown",
  shadowMode: "unknown",
});

function round(value) {
  return Number(Number(value).toFixed(6));
}

function normalizedBounds(value) {
  if (value == null) return null;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("subjectBounds must be an object or null.");
  const result = {};
  for (const key of ["x", "y", "width", "height"]) {
    const number = Number(value[key]);
    if (!Number.isFinite(number)) throw new TypeError(`subjectBounds.${key} must be finite.`);
    result[key] = round(number);
  }
  if (result.x < 0 || result.y < 0 || result.width <= 0 || result.height <= 0 || result.x + result.width > 1 || result.y + result.height > 1) {
    throw new TypeError("subjectBounds must be normalized to a positive rectangle inside 0..1.");
  }
  return Object.freeze(result);
}

export function normalizeImageVisualProfile(value) {
  if (value == null) return DEFAULT_PROFILE;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("visualProfile must be an object or null.");
  const alphaPresent = value.alphaPresent == null ? null : value.alphaPresent;
  if (alphaPresent !== null && typeof alphaPresent !== "boolean") throw new TypeError("alphaPresent must be true, false, or null.");
  const edgeQuality = value.edgeQuality == null ? "unknown" : String(value.edgeQuality).trim().toLowerCase();
  if (!IMAGE_EDGE_QUALITIES.includes(edgeQuality)) throw new TypeError(`edgeQuality must be one of ${IMAGE_EDGE_QUALITIES.join(", ")}.`);
  const shadowMode = value.shadowMode == null ? "unknown" : String(value.shadowMode).trim().toLowerCase();
  if (!IMAGE_SHADOW_MODES.includes(shadowMode)) throw new TypeError(`shadowMode must be one of ${IMAGE_SHADOW_MODES.join(", ")}.`);
  return Object.freeze({
    alphaPresent,
    subjectBounds: normalizedBounds(value.subjectBounds),
    edgeQuality,
    shadowMode,
  });
}

/**
 * Merge byte-derived facts with an optional human/provider declaration. A
 * declaration may add detail that bytes cannot prove (for example a baked
 * shadow), but it cannot contradict a known alpha channel.
 */
export function mergeImageVisualProfiles(derived, declared) {
  const base = normalizeImageVisualProfile(derived);
  const extra = normalizeImageVisualProfile(declared);
  if (base.alphaPresent !== null && extra.alphaPresent !== null && base.alphaPresent !== extra.alphaPresent) {
    throw new TypeError("Declared alphaPresent contradicts the image bytes.");
  }
  return normalizeImageVisualProfile({
    alphaPresent: extra.alphaPresent ?? base.alphaPresent,
    subjectBounds: extra.subjectBounds ?? base.subjectBounds,
    edgeQuality: extra.edgeQuality !== "unknown" ? extra.edgeQuality : base.edgeQuality,
    shadowMode: extra.shadowMode !== "unknown" ? extra.shadowMode : base.shadowMode,
  });
}

function pngHeader(bytes) {
  if (!isPngBytes(bytes)) return null;
  let offset = 8;
  let header;
  let hasTransparencyChunk = false;
  while (offset + 12 <= bytes.length) {
    const length = bytes.readUInt32BE(offset);
    const type = bytes.toString("ascii", offset + 4, offset + 8);
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    if (dataEnd + 4 > bytes.length) return null;
    if (type === "IHDR" && length >= 13) {
      header = {
        width: bytes.readUInt32BE(dataStart),
        height: bytes.readUInt32BE(dataStart + 4),
        bitDepth: bytes[dataStart + 8],
        colorType: bytes[dataStart + 9],
        interlace: bytes[dataStart + 12],
      };
    } else if (type === "tRNS") {
      hasTransparencyChunk = true;
    } else if (type === "IEND") {
      break;
    }
    offset = dataEnd + 4;
  }
  return header ? { ...header, hasTransparencyChunk } : null;
}

function gifHasTransparency(bytes) {
  if (bytes.length < 13 || !["GIF87a", "GIF89a"].includes(bytes.toString("ascii", 0, 6))) return false;
  for (let offset = 13; offset + 8 <= bytes.length; offset += 1) {
    if (bytes[offset] === 0x21 && bytes[offset + 1] === 0xf9 && bytes[offset + 2] === 0x04) {
      return (bytes[offset + 3] & 0x01) === 1;
    }
  }
  return false;
}

function alphaProfileFromPng(bytes, header, maxProfilePixels) {
  if (![4, 6].includes(header.colorType) || header.bitDepth !== 8 || header.interlace !== 0) return {};
  if (header.width * header.height > maxProfilePixels) return {};
  try {
    const decoded = decodePngRgba(bytes);
    let minX = header.width;
    let minY = header.height;
    let maxX = -1;
    let maxY = -1;
    let hasTransparent = false;
    let hasPartial = false;
    for (let y = 0; y < decoded.height; y += 1) {
      for (let x = 0; x < decoded.width; x += 1) {
        const alpha = decoded.pixels[(y * decoded.width + x) * 4 + 3];
        if (alpha < 255) hasTransparent = true;
        if (alpha > 0) {
          minX = Math.min(minX, x);
          minY = Math.min(minY, y);
          maxX = Math.max(maxX, x);
          maxY = Math.max(maxY, y);
        }
        if (alpha > 0 && alpha < 255) hasPartial = true;
      }
    }
    if (!hasTransparent || maxX < 0) return {};
    return {
      subjectBounds: {
        x: minX / decoded.width,
        y: minY / decoded.height,
        width: (maxX - minX + 1) / decoded.width,
        height: (maxY - minY + 1) / decoded.height,
      },
      edgeQuality: hasPartial ? "soft" : "clean",
    };
  } catch {
    return {};
  }
}

/**
 * Derive only conservative facts from image bytes. Subject bounds are emitted
 * for bounded non-interlaced 8-bit PNG alpha masks; JPEGs prove no alpha, and
 * more complex formats remain unknown rather than being guessed.
 */
export function deriveImageVisualProfile(value, { mimeType, maxProfilePixels = 8_000_000 } = {}) {
  const bytes = Buffer.from(value || []);
  const normalizedMime = String(mimeType || "").toLowerCase();
  let profile = { ...DEFAULT_PROFILE };
  if (normalizedMime === "image/jpeg") {
    profile.alphaPresent = false;
  } else {
    const header = pngHeader(bytes);
    if (header) {
      profile.alphaPresent = [4, 6].includes(header.colorType) || header.hasTransparencyChunk;
      if (profile.alphaPresent) profile = { ...profile, ...alphaProfileFromPng(bytes, header, maxProfilePixels) };
    } else if (normalizedMime === "image/gif") {
      profile.alphaPresent = gifHasTransparency(bytes);
    }
  }
  return normalizeImageVisualProfile(profile);
}
