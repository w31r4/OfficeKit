import { createHash } from "node:crypto";

import { decodeSvgDataUrl, svgSourceSafety } from "./svg-text.mjs";

const MAX_SVG_LEAVES = 16_384;
const RGB = /^#(?:[0-9a-f]{3}|[0-9a-f]{6})$/iu;
const NUMBER = /[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/gdu;
const START_TAG = /<(?<prefix>[A-Za-z_][\w.-]*:)?(?<name>[A-Za-z_][\w.-]*)\b(?<attributes>[^<>]*?)\/?>/gdu;
const ATTRIBUTE = /(?<name>[A-Za-z_][\w:.-]*)\s*=\s*(?<quote>["'])(?<value>.*?)\k<quote>/gdu;
const TRANSFORM = /^\s*(?<name>translate|scale|rotate)\s*\((?<args>[\s\S]*)\)\s*$/idu;

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function svgLeafError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function styleSafety(source) {
  const blocked = svgSourceSafety(source);
  if (blocked) return blocked;
  if (/<(?:[A-Za-z_][\w.-]*:)?style\b/iu.test(source) || /\bclass\s*=/iu.test(source)) {
    return "SVG uses stylesheet or class-based styling";
  }
  for (const tag of source.matchAll(START_TAG)) {
    const style = attributesForTag(tag).find((attribute) => attribute.name.toLowerCase() === "style");
    if (!style) continue;
    if (/@import|expression\s*\(/iu.test(style.value)) return "SVG inline style contains active or external CSS";
    for (const match of style.value.matchAll(/url\s*\(\s*(['"]?)(.*?)\1\s*\)/giu)) {
      const target = match[2].trim();
      if (!target.startsWith("#") && !/^data:image\//iu.test(target)) {
        return "SVG inline style contains active or external CSS";
      }
    }
  }
  return "";
}

function normalizedRgb(value) {
  if (!RGB.test(value)) return undefined;
  const hex = value.slice(1);
  return `#${(hex.length === 3 ? [...hex].map((digit) => digit.repeat(2)).join("") : hex).toUpperCase()}`;
}

function normalizedOpacity(value) {
  const raw = typeof value === "number" ? value : Number(String(value).trim());
  if (!Number.isFinite(raw) || raw < 0 || raw > 1) return undefined;
  return raw;
}

function attributesForTag(tag) {
  const source = tag.groups.attributes || "";
  const sourceStart = tag.indices.groups.attributes[0];
  const attributes = [];
  for (const match of source.matchAll(ATTRIBUTE)) {
    attributes.push({
      name: match.groups.name,
      value: match.groups.value,
      sourceStart: sourceStart + match.indices.groups.value[0],
      sourceEnd: sourceStart + match.indices.groups.value[1],
    });
  }
  return attributes;
}

function transformScalars(attribute) {
  const transform = TRANSFORM.exec(attribute.value);
  if (!transform) return [];
  const args = transform.groups.args;
  const numberMatches = [...args.matchAll(NUMBER)];
  const arities = { translate: [1, 2], scale: [1, 2], rotate: [1, 3] };
  if (!arities[transform.groups.name].includes(numberMatches.length) || !validNumberSeparators(args, numberMatches)) return [];
  const values = numberMatches.map((match) => Number(match[0]));
  if (values.some((value) => !Number.isFinite(value))) return [];
  const components = transformComponents(transform.groups.name, values.length);
  const argsStart = transform.indices.groups.args[0];
  return numberMatches.map((match, index) => ({
    component: components[index],
    value: values[index],
    raw: match[0],
    sourceStart: attribute.sourceStart + argsStart + match.indices[0][0],
    sourceEnd: attribute.sourceStart + argsStart + match.indices[0][1],
  }));
}

function validNumberSeparators(args, matches) {
  let cursor = 0;
  for (const [index, match] of matches.entries()) {
    const gap = args.slice(cursor, match.indices[0][0]);
    if (index === 0 ? !/^\s*$/u.test(gap) : !/^(?:\s*,\s*|\s+)$/u.test(gap)) return false;
    cursor = match.indices[0][1];
  }
  return /^\s*$/u.test(args.slice(cursor));
}

function transformComponents(name, count) {
  if (name === "translate") return count === 1 ? ["x"] : ["x", "y"];
  if (name === "scale") return count === 1 ? ["scale"] : ["scaleX", "scaleY"];
  return count === 1 ? ["angle"] : ["angle", "centerX", "centerY"];
}

function normalizeTransformScalar(component, value) {
  const number = typeof value === "number" ? value : Number(String(value).trim());
  if (!Number.isFinite(number)) return undefined;
  if (["scale", "scaleX", "scaleY"].includes(component)) {
    if (Math.abs(number) < 0.001 || Math.abs(number) > 1_000) return undefined;
  } else if (component === "angle") {
    if (Math.abs(number) > 360_000) return undefined;
  } else if (Math.abs(number) > 1_000_000) {
    return undefined;
  }
  return number;
}

function createLeaf({ sourceSha256, sourceRevisionSha256, scopeId, tag, tagIndex, attribute, leafKind, value, raw, component, sourceStart, sourceEnd }) {
  const expectedHash = sha256(raw);
  const seed = `${scopeId}\0${sourceSha256}\0${tagIndex}\0${tag}\0${attribute}\0${component || ""}\0${sourceStart}\0${expectedHash}`;
  return {
    record: Object.freeze({
      kind: "svgLeaf",
      leafKind,
      id: `sl_${sha256(seed).slice(0, 32)}`,
      tag,
      attribute,
      ...(component ? { component } : {}),
      value,
      expectedHash,
      sourceSha256,
      ...(sourceRevisionSha256 ? { sourceRevisionSha256 } : {}),
    }),
    sourceStart,
    sourceEnd,
    raw,
  };
}

function parseSvgLeaves(decoded, scope = {}) {
  const blockedReason = styleSafety(decoded.source);
  if (blockedReason) return { blockedReason, leaves: [] };
  const sourceSha256 = sha256(decoded.bytes);
  const leaves = [];
  let tagIndex = 0;
  for (const tag of decoded.source.matchAll(START_TAG)) {
    tagIndex += 1;
    const attributes = attributesForTag(tag);
    if (attributes.some((attribute) => ["style", "class"].includes(attribute.name.toLowerCase()))) continue;
    for (const attribute of attributes) {
      const name = attribute.name.toLowerCase();
      if (name === "fill" || name === "stroke") {
        const value = normalizedRgb(attribute.value);
        if (value) leaves.push(createLeaf({
          sourceSha256,
          sourceRevisionSha256: scope.sourceRevisionSha256,
          scopeId: scope.scopeId || "",
          tag: tag.groups.name,
          tagIndex,
          attribute: name,
          leafKind: name === "fill" ? "svgFillRgb" : "svgStrokeRgb",
          value,
          raw: attribute.value,
          sourceStart: attribute.sourceStart,
          sourceEnd: attribute.sourceEnd,
        }));
      } else if (["opacity", "fill-opacity", "stroke-opacity"].includes(name)) {
        const value = normalizedOpacity(attribute.value);
        if (value !== undefined) leaves.push(createLeaf({
          sourceSha256,
          sourceRevisionSha256: scope.sourceRevisionSha256,
          scopeId: scope.scopeId || "",
          tag: tag.groups.name,
          tagIndex,
          attribute: name,
          leafKind: "svgOpacity",
          value,
          raw: attribute.value,
          sourceStart: attribute.sourceStart,
          sourceEnd: attribute.sourceEnd,
        }));
      } else if (name === "transform") {
        for (const scalar of transformScalars(attribute)) leaves.push(createLeaf({
          sourceSha256,
          sourceRevisionSha256: scope.sourceRevisionSha256,
          scopeId: scope.scopeId || "",
          tag: tag.groups.name,
          tagIndex,
          attribute: name,
          leafKind: "svgTransformScalar",
          value: scalar.value,
          raw: scalar.raw,
          component: scalar.component,
          sourceStart: scalar.sourceStart,
          sourceEnd: scalar.sourceEnd,
        }));
      }
      if (leaves.length > MAX_SVG_LEAVES) return { blockedReason: "SVG editable leaf budget exceeded", leaves: [] };
    }
  }
  return { blockedReason: "", leaves };
}

export function inspectSvgLeaves(dataUrl, scope = {}) {
  const decoded = decodeSvgDataUrl(dataUrl);
  if (!decoded) return Object.freeze({ supported: false, reason: "image is not a bounded base64 SVG" });
  const parsed = parseSvgLeaves(decoded, scope);
  const sourceSha256 = sha256(decoded.bytes);
  if (parsed.blockedReason) return Object.freeze({ supported: false, reason: parsed.blockedReason, sourceSha256 });
  return Object.freeze({
    supported: parsed.leaves.length > 0,
    reason: parsed.leaves.length ? "" : "SVG contains no directly editable style or transform leaves",
    sourceSha256,
    ...(scope.sourceRevisionSha256 ? { sourceRevisionSha256: scope.sourceRevisionSha256 } : {}),
    leaves: Object.freeze(parsed.leaves.map((leaf) => leaf.record)),
  });
}

export function editSvgLeaf(dataUrl, leafId, update = {}, scope = {}) {
  const decoded = decodeSvgDataUrl(dataUrl);
  if (!decoded) throw new TypeError("SVG leaf editing requires a bounded base64 SVG image.");
  const parsed = parseSvgLeaves(decoded, scope);
  if (parsed.blockedReason) {
    throw svgLeafError("unsupported_presentation_svg_leaf", `SVG leaf editing is blocked: ${parsed.blockedReason}.`);
  }
  if (!update || typeof update !== "object" || Array.isArray(update) ||
      Object.keys(update).sort().join(",") !== "expectedHash,value") {
    throw new TypeError("SVG leaf edit accepts exactly expectedHash and value.");
  }
  const leaf = parsed.leaves.find((candidate) => candidate.record.id === leafId);
  if (!leaf) throw svgLeafError("presentation_svg_leaf_not_issued", "SVG leaf target was not issued for the current image bytes.");
  if (String(update.expectedHash || "").toLowerCase() !== leaf.record.expectedHash) {
    throw svgLeafError("presentation_svg_leaf_stale", "SVG leaf expectedHash does not match the current image bytes.");
  }
  let value;
  if (leaf.record.leafKind === "svgFillRgb" || leaf.record.leafKind === "svgStrokeRgb") value = normalizedRgb(String(update.value || ""));
  else if (leaf.record.leafKind === "svgOpacity") value = normalizedOpacity(update.value);
  else value = normalizeTransformScalar(leaf.record.component, update.value);
  if (value === undefined) throw svgLeafError("invalid_presentation_svg_leaf", `Invalid ${leaf.record.leafKind} value.`);
  if (value === leaf.record.value) throw svgLeafError("presentation_svg_leaf_noop", "SVG leaf edit must change its source value.");
  const replacement = typeof value === "number" ? String(value) : value;
  const source = `${decoded.source.slice(0, leaf.sourceStart)}${replacement}${decoded.source.slice(leaf.sourceEnd)}`;
  return `data:image/svg+xml;base64,${Buffer.from(new TextEncoder().encode(source)).toString("base64")}`;
}
