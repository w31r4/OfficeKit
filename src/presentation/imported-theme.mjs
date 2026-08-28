import { createHash } from "node:crypto";

import { loadOoxmlZipWithinBudget, ooxmlResolveRelationshipTarget, ooxmlSafePartPath } from "../ooxml/package.mjs";
import { decoder } from "../shared/binary.mjs";

const THEME_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const THEME_RELATIONSHIP_ALT = "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";
const COLOR_ROLES = Object.freeze([
  ["dk1", "tx1"], ["lt1", "bg1"], ["dk2", "tx2"], ["lt2", "bg2"],
  ["accent1", "accent1"], ["accent2", "accent2"], ["accent3", "accent3"],
  ["accent4", "accent4"], ["accent5", "accent5"], ["accent6", "accent6"],
  ["hlink", "hlink"], ["folHlink", "folHlink"],
]);

/**
 * Read the source theme as bounded design evidence. This intentionally does
 * not turn the theme into an editable authoring theme: source-bound exports
 * keep the original theme opaque, while designProfile can still describe the
 * visual language an Agent should preserve when continuing the deck.
 */
export async function parseImportedPresentationTheme(bytes, limits = {}) {
  try {
    const zip = await loadOoxmlZipWithinBudget(bytes, limits, "PPTX theme profile");
    const partPath = await findThemePartPath(zip);
    if (!partPath) return undefined;
    const entry = zip.file(partPath);
    if (!entry) return undefined;
    const xmlBytes = await entry.async("uint8array");
    const xml = decoder.decode(xmlBytes);
    const profile = themeProfile(xml, partPath, xmlBytes);
    return profile && Object.keys(profile.colors).length + Object.keys(profile.fonts).length
      ? deepFreeze(profile)
      : undefined;
  } catch {
    // Theme evidence is optional. A malformed or non-standard theme must not
    // make an otherwise valid source-bound import fail; the original bytes
    // remain preserved and the design profile reports its semantic projection.
    return undefined;
  }
}

async function findThemePartPath(zip) {
  const masterPaths = Object.keys(zip.files)
    .filter((name) => /^ppt\/slideMasters\/slide\d+\.xml$/u.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  for (const masterPath of masterPaths) {
    const relsPath = relationshipPartPath(masterPath);
    const rels = zip.file(relsPath);
    if (!rels) continue;
    const relsXml = decoder.decode(await rels.async("uint8array"));
    for (const relationship of relationshipEntries(relsXml)) {
      if (relationship.type !== THEME_RELATIONSHIP && relationship.type !== THEME_RELATIONSHIP_ALT) continue;
      try {
        const target = ooxmlSafePartPath(ooxmlResolveRelationshipTarget(masterPath, relationship.target), "PPTX theme");
        if (zip.file(target)) return target;
      } catch {
        // Ignore an unsafe optional relationship and continue looking for the
        // next canonical theme relation.
      }
    }
  }
  return Object.keys(zip.files)
    .filter((name) => /^ppt\/theme\/theme\d+\.xml$/u.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))[0];
}

function relationshipPartPath(sourcePath) {
  const slash = sourcePath.lastIndexOf("/");
  const directory = slash >= 0 ? sourcePath.slice(0, slash) : "";
  const name = slash >= 0 ? sourcePath.slice(slash + 1) : sourcePath;
  return `${directory}/_rels/${name}.rels`;
}

function relationshipEntries(xml) {
  return [...String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b([^>]*)\/?\s*>/gu)]
    .map((match) => {
      const attrs = attributes(match[1]);
      return { type: attrs.Type || "", target: attrs.Target || "", targetMode: attrs.TargetMode || "" };
    })
    .filter((relationship) => relationship.type && relationship.target && !/^external$/iu.test(relationship.targetMode || ""));
}

function themeProfile(xml, partPath, xmlBytes) {
  const themeAttrs = attributes(openingTag(xml, "theme"));
  const colorScheme = block(xml, "clrScheme");
  const fontScheme = block(xml, "fontScheme");
  const colors = {};
  for (const [rawRole, role] of COLOR_ROLES) {
    const colorBlock = block(colorScheme, rawRole);
    const color = colorValue(colorBlock);
    if (color) colors[role] = color;
  }
  const fonts = {};
  const major = block(fontScheme, "majorFont");
  const minor = block(fontScheme, "minorFont");
  addFont(fonts, "major", major, "latin");
  addFont(fonts, "minor", minor, "latin");
  addFont(fonts, "majorEastAsia", major, "ea");
  addFont(fonts, "minorEastAsia", minor, "ea");
  addFont(fonts, "majorComplexScript", major, "cs");
  addFont(fonts, "minorComplexScript", minor, "cs");
  if (!colors.tx1 && !colors.bg1 && !fonts.major && !fonts.minor) return undefined;
  const xmlSha256 = createHash("sha256").update(xmlBytes).digest("hex");
  return {
    kind: "theme",
    id: `theme/source/${xmlSha256.slice(0, 16)}`,
    name: decodeXml(themeAttrs.name || "Imported theme"),
    colorSchemeName: decodeXml(attributes(openingTag(colorScheme, "clrScheme")).name || ""),
    colors,
    fonts,
    textStyles: {},
    colorMap: {
      bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2",
      accent1: "accent1", accent2: "accent2", accent3: "accent3",
      accent4: "accent4", accent5: "accent5", accent6: "accent6",
      hlink: "hlink", folHlink: "folHlink",
    },
    source: {
      sourceBound: true,
      editable: false,
      partPath,
      xmlSha256,
    },
  };
}

function addFont(fonts, key, parent, child) {
  const value = decodeXml(attributes(openingTag(parent, child)).typeface || "");
  if (value) fonts[key] = value;
}

function colorValue(source) {
  if (!source) return undefined;
  const srgb = attributes(openingTag(source, "srgbClr")).val;
  if (/^[0-9a-f]{6}$/iu.test(srgb || "")) return `#${srgb.toLowerCase()}`;
  const system = attributes(openingTag(source, "sysClr"));
  if (/^[0-9a-f]{6}$/iu.test(system.lastClr || "")) return `#${system.lastClr.toLowerCase()}`;
  return undefined;
}

function block(source, localName) {
  const match = new RegExp(`<(?:(?:[A-Za-z_][\\w.-]*):)?${localName}\\b[^>]*>([\\s\\S]*?)<\\/(?:(?:[A-Za-z_][\\w.-]*):)?${localName}>`, "iu").exec(String(source || ""));
  return match ? match[0] : "";
}

function openingTag(source, localName) {
  const match = new RegExp(`<(?:(?:[A-Za-z_][\\w.-]*):)?${localName}\\b([^>]*)>`, "iu").exec(String(source || ""));
  return match ? match[1] : "";
}

function attributes(source) {
  const result = {};
  for (const match of String(source || "").matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"]*)"|'([^']*)')/gu)) {
    result[match[1].split(":").pop()] = match[2] ?? match[3] ?? "";
  }
  return result;
}

function decodeXml(value) {
  return String(value || "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function deepFreeze(value) {
  if (!value || typeof value !== "object" || Object.isFrozen(value)) return value;
  Object.freeze(value);
  for (const child of Object.values(value)) deepFreeze(child);
  return value;
}
