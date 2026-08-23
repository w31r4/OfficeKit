import { createHash } from "node:crypto";

import { decodeXml } from "../ooxml/source-reference-xml.mjs";
import { xmlEscape } from "../shared/xml.mjs";

const SVG_DATA_URL = /^data:image\/svg\+xml;base64,([A-Za-z0-9+/=\s]+)$/iu;
const SVG_TEXT_TAGS = Object.freeze([
  /<(?<prefix>[A-Za-z_][\w.-]*:)?(?<name>text)\b(?<attributes>[^>]*)>(?<value>[\s\S]*?)<\/\k<prefix>(?:\k<name>)\s*>/giu,
  /<(?<prefix>[A-Za-z_][\w.-]*:)?(?<name>tspan)\b(?<attributes>[^>]*)>(?<value>[\s\S]*?)<\/\k<prefix>(?:\k<name>)\s*>/giu,
]);
const MAX_SVG_TEXT_BYTES = 16 * 1024 * 1024;
const MAX_SVG_TEXT_LENGTH = 32_767;
const MAX_SVG_TEXT_NODES = 4_096;

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

export function decodeSvgDataUrl(dataUrl) {
  const match = SVG_DATA_URL.exec(String(dataUrl || ""));
  if (!match) return undefined;
  const base64 = match[1].replace(/\s+/gu, "");
  const bytes = Buffer.from(base64, "base64");
  if (!bytes.length || bytes.length > MAX_SVG_TEXT_BYTES) return undefined;
  let source;
  try {
    source = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return undefined;
  }
  if (!/^\s*<svg\b[^>]*>/iu.test(source) || !/<\/svg>\s*$/iu.test(source)) return undefined;
  return { bytes, source };
}

export function svgSourceSafety(source) {
  if (/<\s*(?:script|foreignObject|iframe|object|embed)\b/iu.test(source) ||
      /<!\s*(?:DOCTYPE|ENTITY)\b/iu.test(source) ||
      /\bon[A-Za-z][\w.-]*\s*=/iu.test(source) ||
      /(?:href|xlink:href)\s*=\s*(['"])(?!#|data:image\/)[^'"]+\1/iu.test(source)) {
    return "SVG contains active content or an external reference";
  }
  return "";
}

function validSvgText(value) {
  return value.length <= MAX_SVG_TEXT_LENGTH &&
    !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value);
}

function parseSvgTextNodes(source) {
  const nodes = [];
  const matches = [];
  for (const pattern of SVG_TEXT_TAGS) {
    for (const match of source.matchAll(pattern)) matches.push(match);
  }
  matches.sort((left, right) => (left.index ?? 0) - (right.index ?? 0));
  for (const match of matches) {
    const value = match.groups?.value || "";
    // A parent <text> containing tspans or other markup is not itself a safe
    // replacement target. Its direct tspans are still exposed below.
    if (/<[A-Za-z_][\w:.-]*\b/iu.test(value) || !validSvgText(decodeXml(value))) continue;
    if (nodes.length >= MAX_SVG_TEXT_NODES) return { nodes: [], reason: "SVG text node budget exceeded" };
    const start = match.index ?? -1;
    const openEnd = start + (match[0].indexOf(">") + 1);
    const valueStart = openEnd;
    const valueEnd = valueStart + value.length;
    const text = decodeXml(value);
    const index = nodes.length;
    nodes.push({
      id: `svg-text-${index + 1}`,
      index,
      tag: match.groups?.name || "text",
      text,
      expectedHash: sha256(text),
      sourceStart: valueStart,
      sourceEnd: valueEnd,
    });
  }
  return { nodes, reason: "" };
}

export function inspectSvgText(dataUrl) {
  const decoded = decodeSvgDataUrl(dataUrl);
  if (!decoded) return Object.freeze({ supported: false, reason: "image is not a bounded base64 SVG" });
  const blockedReason = svgSourceSafety(decoded.source);
  if (blockedReason) return Object.freeze({ supported: false, reason: blockedReason, sourceSha256: sha256(decoded.bytes) });
  const parsed = parseSvgTextNodes(decoded.source);
  if (parsed.reason) return Object.freeze({ supported: false, reason: parsed.reason, sourceSha256: sha256(decoded.bytes) });
  return Object.freeze({
    supported: parsed.nodes.length > 0,
    reason: parsed.nodes.length ? "" : "SVG contains no directly editable text nodes",
    sourceSha256: sha256(decoded.bytes),
    nodes: Object.freeze(parsed.nodes.map(({ id, index, tag, text, expectedHash }) => Object.freeze({ id, index, tag, text, expectedHash }))),
  });
}

export function editSvgText(dataUrl, nodeId, update = {}) {
  const decoded = decodeSvgDataUrl(dataUrl);
  if (!decoded) throw new TypeError("SVG text editing requires a bounded base64 SVG image.");
  const blockedReason = svgSourceSafety(decoded.source);
  if (blockedReason) {
    const error = new Error(`SVG text editing is blocked: ${blockedReason}.`);
    error.code = "unsupported_presentation_svg_text";
    throw error;
  }
  const parsed = parseSvgTextNodes(decoded.source);
  if (parsed.reason) {
    const error = new Error(`SVG text editing is blocked: ${parsed.reason}.`);
    error.code = "unsupported_presentation_svg_text";
    throw error;
  }
  if (!update || typeof update !== "object" || Array.isArray(update) ||
      Object.keys(update).sort().join(",") !== "expectedHash,value") {
    throw new TypeError("SVG text edit accepts exactly expectedHash and value.");
  }
  const node = parsed.nodes.find((candidate) => candidate.id === nodeId);
  if (!node) {
    const error = new Error("SVG text edit target was not issued for the current image bytes.");
    error.code = "presentation_svg_text_not_issued";
    throw error;
  }
  if (String(update.expectedHash || "").toLowerCase() !== node.expectedHash) {
    const error = new Error("SVG text edit expectedHash does not match the current image bytes.");
    error.code = "presentation_svg_text_stale";
    throw error;
  }
  const value = String(update.value ?? "");
  if (!validSvgText(value)) {
    const error = new Error("SVG text edit value contains controls or exceeds 32767 characters.");
    error.code = "invalid_presentation_svg_text";
    throw error;
  }
  if (value === node.text) {
    const error = new Error("SVG text edit must change its source value.");
    error.code = "presentation_svg_text_noop";
    throw error;
  }
  const source = `${decoded.source.slice(0, node.sourceStart)}${xmlEscape(value)}${decoded.source.slice(node.sourceEnd)}`;
  const bytes = new TextEncoder().encode(source);
  return `data:image/svg+xml;base64,${Buffer.from(bytes).toString("base64")}`;
}
