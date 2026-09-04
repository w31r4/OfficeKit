import { toUint8Array } from "./binary.mjs";
import { deriveImageVisualProfile } from "./image-profile.mjs";

const PNG_SIGNATURE = Buffer.from("89504e470d0a1a0a", "hex");
const JPEG_SOF_MARKERS = new Set([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf]);
const GIF_SIGNATURES = new Set(["GIF87a", "GIF89a"]);
const DEFAULT_MAX_DIMENSION = 10_000_000;

export const IMAGE_MIME_TYPES = Object.freeze([
  "image/png",
  "image/jpeg",
  "image/gif",
  "image/svg+xml",
]);

export function normalizeImageMimeType(value) {
  const normalized = String(value || "").split(";", 1)[0].trim().toLowerCase();
  return normalized === "image/jpg" ? "image/jpeg" : normalized;
}

export function imageExtensionForMimeType(value) {
  const mimeType = normalizeImageMimeType(value);
  if (mimeType === "image/png") return "png";
  if (mimeType === "image/jpeg") return "jpg";
  if (mimeType === "image/gif") return "gif";
  if (mimeType === "image/svg+xml") return "svg";
  return undefined;
}

function boundedDimensions(width, height, label, maxDimension) {
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0 || width > maxDimension || height > maxDimension) {
    throw new TypeError(`${label} does not expose bounded positive intrinsic dimensions.`);
  }
  return { width, height };
}

function pngDimensions(bytes, label, maxDimension) {
  if (bytes.length < 24 || !bytes.subarray(0, 8).equals(PNG_SIGNATURE)) return undefined;
  return boundedDimensions(bytes.readUInt32BE(16), bytes.readUInt32BE(20), label, maxDimension);
}

function gifDimensions(bytes, label, maxDimension) {
  if (bytes.length < 10 || !GIF_SIGNATURES.has(bytes.subarray(0, 6).toString("ascii"))) return undefined;
  return boundedDimensions(bytes.readUInt16LE(6), bytes.readUInt16LE(8), label, maxDimension);
}

function jpegDimensions(bytes, label, maxDimension) {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined;
  let offset = 2;
  while (offset + 3 < bytes.length) {
    while (offset < bytes.length && bytes[offset] !== 0xff) offset += 1;
    while (offset < bytes.length && bytes[offset] === 0xff) offset += 1;
    if (offset >= bytes.length) break;
    const marker = bytes[offset++];
    if (marker === 0xd8 || marker === 0xd9 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (offset + 2 > bytes.length) break;
    const length = bytes.readUInt16BE(offset);
    if (length < 2 || offset + length > bytes.length) break;
    if (JPEG_SOF_MARKERS.has(marker)) {
      if (length < 7) break;
      return boundedDimensions(bytes.readUInt16BE(offset + 3), bytes.readUInt16BE(offset + 5), label, maxDimension);
    }
    offset += length;
  }
  throw new TypeError(`${label} does not expose intrinsic dimensions in a supported JPEG SOF segment.`);
}

function svgSource(bytes) {
  return bytes.toString("utf8").replace(/^\uFEFF/, "");
}

export function isSafeSvgImageBytes(value) {
  const bytes = Buffer.from(toUint8Array(value));
  const source = svgSource(bytes);
  if (!/^\s*(?:<\?xml[\s\S]*?\?>\s*)?(?:(?:<!--[\s\S]*?-->\s*)*)<svg(?:\s|>)/i.test(source)) return false;
  if (/<!DOCTYPE|<!ENTITY|<\?xml-stylesheet|<\s*(?:script|foreignObject)\b|\son[a-z]+\s*=|@import\b/i.test(source)) return false;
  for (const match of source.matchAll(/\s(?:href|xlink:href)\s*=\s*(["'])(.*?)\1/gi)) {
    const target = match[2].trim();
    if (target && !target.startsWith("#") && !/^data:image\/(?:png|jpe?g|gif);base64,/i.test(target)) return false;
  }
  for (const match of source.matchAll(/url\(\s*(["']?)(.*?)\1\s*\)/gi)) {
    const target = match[2].trim();
    if (target && !target.startsWith("#") && !/^data:image\/(?:png|jpe?g|gif);base64,/i.test(target)) return false;
  }
  return true;
}

function svgDimensions(bytes, label, maxDimension) {
  const source = svgSource(bytes);
  const root = /<svg\b([^>]*)>/i.exec(source)?.[1];
  if (root == null) return undefined;
  const numericAttribute = (name) => {
    const match = new RegExp(`\\b${name}\\s*=\\s*(["'])([+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+))(?:px)?\\1`, "i").exec(root);
    return match ? Number(match[2]) : undefined;
  };
  const width = numericAttribute("width");
  const height = numericAttribute("height");
  if (Number.isFinite(width) && Number.isFinite(height)) return boundedDimensions(width, height, label, maxDimension);
  const viewBox = /\bviewBox\s*=\s*(["'])([-+0-9.eE\s,]+)\1/i.exec(root)?.[2]
    ?.trim().split(/[\s,]+/).map(Number);
  if (viewBox?.length === 4) return boundedDimensions(viewBox[2], viewBox[3], label, maxDimension);
  throw new TypeError(`${label} requires bounded numeric width/height or viewBox dimensions.`);
}

export function inspectImageBytes(value, options = {}) {
  const label = String(options.label || "Image");
  const bytes = Buffer.from(toUint8Array(value));
  const maxBytes = Number(options.maxBytes ?? Number.POSITIVE_INFINITY);
  const maxPixels = Number(options.maxPixels ?? Number.POSITIVE_INFINITY);
  const maxDimension = Number(options.maxDimension ?? DEFAULT_MAX_DIMENSION);
  for (const [name, limit] of [["maxBytes", maxBytes], ["maxPixels", maxPixels], ["maxDimension", maxDimension]]) {
    if (!(limit === Number.POSITIVE_INFINITY || (Number.isFinite(limit) && limit > 0))) {
      throw new TypeError(`Image inspection ${name} must be a positive finite value or Infinity.`);
    }
  }
  if (!bytes.length || bytes.length > maxBytes) throw new RangeError(`${label} must contain 1 through ${maxBytes} bytes.`);

  let mimeType;
  let size;
  if ((size = pngDimensions(bytes, label, maxDimension))) mimeType = "image/png";
  else if ((size = jpegDimensions(bytes, label, maxDimension))) mimeType = "image/jpeg";
  else if ((size = gifDimensions(bytes, label, maxDimension))) mimeType = "image/gif";
  else if (options.allowSvg !== false && /^\s*(?:<\?xml[\s\S]*?\?>\s*)?(?:(?:<!--[\s\S]*?-->\s*)*)<svg(?:\s|>)/i.test(svgSource(bytes))) {
    if (!isSafeSvgImageBytes(bytes)) throw new TypeError(`${label} must be a safe SVG without scripts, events, external references, DTDs, or foreignObject.`);
    size = svgDimensions(bytes, label, maxDimension);
    mimeType = "image/svg+xml";
  } else {
    throw new TypeError(`${label} must be a PNG, JPEG, GIF${options.allowSvg === false ? "" : ", or safe SVG"} image.`);
  }

  const declaredMimeType = normalizeImageMimeType(options.declaredMimeType);
  if (declaredMimeType && declaredMimeType !== mimeType) {
    throw new TypeError(`${label} bytes use ${mimeType}, not declared content type ${declaredMimeType}.`);
  }
  const pixels = size.width * size.height;
  if (pixels > maxPixels) throw new RangeError(`${label} exceeds the ${maxPixels}-pixel limit.`);
  return Object.freeze({
    mimeType,
    extension: imageExtensionForMimeType(mimeType),
    width: size.width,
    height: size.height,
    pixels,
    byteLength: bytes.length,
    visualProfile: deriveImageVisualProfile(bytes, { mimeType }),
  });
}

export function imageDataUrlFromBytes(value, mimeType) {
  const bytes = Buffer.from(toUint8Array(value));
  const normalized = normalizeImageMimeType(mimeType);
  inspectImageBytes(bytes, { declaredMimeType: normalized, label: "Image blob" });
  return `data:${normalized};base64,${bytes.toString("base64")}`;
}
