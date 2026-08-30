import { createHash } from "node:crypto";

import { decodeXml } from "../ooxml/source-reference-xml.mjs";
import { toUint8Array } from "../shared/binary.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { aid } from "../shared/ids.mjs";
import { attrEscape, xmlEscape } from "../shared/xml.mjs";
import { presentationElementDeletionCapability } from "./element-deletion.mjs";
import { directPresentationChildren } from "./group-shapes.mjs";

const MAX_EMBEDDED_WORKBOOK_BYTES = 16 * 1024 * 1024;
const MAX_EMBEDDED_OFFICE_PACKAGE_BYTES = 16 * 1024 * 1024;
const MAX_DIAGRAM_NODE_TEXT_LENGTH = 32_767;
const MAX_NATIVE_TEXT_LENGTH = 32_767;
const MAX_NATIVE_LINE_WIDTH_EMU = 20_116_800;
const MAX_NATIVE_STYLE_LEAVES = 4_096;
const MAX_DIAGRAM_NODE_RUNS = 256;
const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const CHART_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const TABLE_GRAPHIC_DATA_URI_PATTERN = /\buri\s*=\s*["']http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/table["']/iu;
const NATIVE_TEXT_TAG = /<(?<prefix>[A-Za-z_][\w.-]*:)?t\b[^>]*?(?:\/\s*>|>(?<value>[^<]*)<\/(?:[A-Za-z_][\w.-]*:)?t\s*>)/giu;
const NATIVE_TEXT_RUN = /<(?<prefix>[A-Za-z_][\w.-]*:)?r\b[^>]*>(?<value>[\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?r\s*>/giu;
const NATIVE_TEXT_CELL = /<(?<prefix>[A-Za-z_][\w.-]*:)?tc\b[^>]*>(?<value>[\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?tc\s*>/giu;
const NATIVE_TEXT_SHAPE = /<(?<prefix>[A-Za-z_][\w.-]*:)?sp\b[^>]*>(?<value>[\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?sp\s*>/giu;
const NATIVE_SPPR_TAG = /<(?<prefix>[A-Za-z_][\w.-]*:)?spPr\b[^>]*>(?<value>[\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?spPr\s*>/giu;
const NATIVE_LINE_TAG = /<(?<prefix>[A-Za-z_][\w.-]*:)?ln\b(?<attributes>[^>]*)>(?<value>[\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?ln\s*>/giu;
const NATIVE_PRESET_DASH_TAG = /<(?<prefix>[A-Za-z_][\w.-]*:)?prstDash\b(?<attributes>[^>]*)\s*\/\s*>/giu;
const NATIVE_SOLID_FILL_TAG = /<(?<prefix>[A-Za-z_][\w.-]*:)?solidFill\b[^>]*>(?<value>[\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?solidFill\s*>/giu;
const NATIVE_COLOR_OPEN_TAG = /<(?<prefix>[A-Za-z_][\w.-]*:)?(?<name>[A-Za-z_][\w.-]*Clr)\b(?<attributes>[^>]*)>/giu;
const NATIVE_SCHEME_COLORS = Object.freeze({
  bg1: "bg1",
  tx1: "tx1",
  bg2: "bg2",
  tx2: "tx2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folhlink: "folHlink",
  dk1: "dk1",
  lt1: "lt1",
  dk2: "dk2",
  lt2: "lt2",
});
const NATIVE_LINE_STYLES = Object.freeze({
  solid: "solid",
  dash: "dashed",
  dashed: "dashed",
  dot: "dotted",
  dotted: "dotted",
  dashDot: "dash-dot",
  "dash-dot": "dash-dot",
  lgDashDotDot: "dash-dot-dot",
  "dash-dot-dot": "dash-dot-dot",
});
const NATIVE_LINE_CAPS = Object.freeze({
  flat: "flat",
  rnd: "round",
  round: "round",
  sq: "square",
  square: "square",
});
const NATIVE_LINE_JOINS = Object.freeze({
  round: "round",
  bevel: "bevel",
  miter: "miter",
});
const NATIVE_LINE_ARROWS = Object.freeze({
  none: "none",
  triangle: "triangle",
  stealth: "stealth",
  diamond: "diamond",
  oval: "oval",
  arrow: "arrow",
});
const DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";

function sha256(value) {
  return createHash("sha256").update(value, "utf8").digest("hex");
}

function deriveNativeTextLeaves(rawXml, nativeKind) {
  const source = String(rawXml || "");
  const table = nativeKind === "graphicFrame" && TABLE_GRAPHIC_DATA_URI_PATTERN.test(source);
  const group = nativeKind === "group";
  if (!table && !group) return undefined;
  const cells = table ? [...source.matchAll(NATIVE_TEXT_CELL)] : [];
  const shapes = group ? [...source.matchAll(NATIVE_TEXT_SHAPE)] : [];
  const runs = [...source.matchAll(NATIVE_TEXT_RUN)];
  const texts = [...source.matchAll(NATIVE_TEXT_TAG)];
  if ((!table && !shapes.length) || !texts.length || texts.length > 4_096) return undefined;
  const inRange = (match, container) => {
    const start = match.index ?? -1;
    const end = start + match[0].length;
    const containerStart = container.index ?? -1;
    return start >= containerStart && end <= containerStart + container[0].length;
  };
  if (texts.some((text) => {
    if (!runs.some((run) => inRange(text, run))) return true;
    if (table) return !cells.some((cell) => inRange(text, cell));
    return cells.some((cell) => inRange(text, cell)) || shapes.filter((shape) => inRange(text, shape)).length !== 1;
  })) return undefined;
  const leaves = texts.map((match, index) => {
    const text = decodeXml(match.groups?.value || "");
    if (!validDiagramNodeText(text)) return undefined;
    return Object.freeze({ textLeafIndex: index, text, expectedHash: sha256(text) });
  });
  return leaves.every(Boolean) ? Object.freeze(leaves) : undefined;
}

function nativeTextRecord(leaves) {
  return leaves ? Object.freeze(leaves.map((leaf) => Object.freeze({
    textLeafIndex: leaf.textLeafIndex,
    text: leaf.text,
    expectedHash: sha256(leaf.text),
  }))) : undefined;
}

function nativeTagAttributes(tag) {
  return [...String(tag || "").matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/gu)]
    .map((match) => ({ name: match[1], value: decodeXml(match[3]) }));
}

function nativeStyleAttributes(tag) {
  return nativeTagAttributes(tag).filter((attribute) => !/^xmlns(?::|$)/iu.test(attribute.name));
}

function hasAllowedNativeNamespaceAttributes(tag) {
  return nativeTagAttributes(tag)
    .filter((attribute) => /^xmlns(?::|$)/iu.test(attribute.name))
    .every((attribute) => attribute.value === DRAWINGML_NAMESPACE);
}

function nativeSchemeColorToken(value) {
  return NATIVE_SCHEME_COLORS[String(value || "").trim().toLowerCase()];
}

function nativeLineStyleToken(value) {
  return NATIVE_LINE_STYLES[String(value || "").trim()];
}

function nativeLineCapToken(value) {
  return NATIVE_LINE_CAPS[String(value || "").trim().toLowerCase()];
}

function nativeLineJoinToken(value) {
  return NATIVE_LINE_JOINS[String(value || "").trim().toLowerCase()];
}

function nativeLineArrowToken(value) {
  return NATIVE_LINE_ARROWS[String(value || "").trim().toLowerCase()];
}

function hasCanonicalStyleReferenceColor(xml) {
  const source = String(xml || "");
  const open = /^<[^>]+>/u.exec(source)?.[0] || "";
  if (!open || !/\/\s*>$/u.test(source)) return false;
  const attributes = nativeStyleAttributes(open);
  const localName = /^<(?:(?:[A-Za-z_][\w.-]*):)?(?<name>[A-Za-z_][\w.-]*)\b/u.exec(open)?.groups?.name?.toLowerCase();
  return hasAllowedNativeNamespaceAttributes(open) && attributes.length === 1 && attributes[0].name.split(":").pop()?.toLowerCase() === "val" && attributes[0].value &&
    (localName === "schemeclr"
      ? Boolean(nativeSchemeColorToken(attributes[0].value))
      : localName === "srgbclr" && /^[0-9a-f]{6}$/iu.test(attributes[0].value));
}

function hasCanonicalStyleReference(node) {
  const xml = String(node?.xml || "");
  const open = /^<[^>]+>/u.exec(xml)?.[0] || "";
  if (!open || /\/\s*>$/u.test(xml) || !hasAllowedNativeNamespaceAttributes(open)) return false;
  const attributes = nativeStyleAttributes(open);
  if (attributes[0].name.split(":").pop()?.toLowerCase() !== "idx") return false;
  const local = String(node?.localName || "").toLowerCase();
  const index = attributes[0].value;
  if (local === "fontref") {
    if (index !== "minor" && index !== "major") return false;
  } else if (!/^(?:0|[1-9][0-9]*)$/u.test(index) || Number(index) > 32) {
    return false;
  }
  const children = directPresentationChildren(xml, node.localName);
  if (children.length !== 1 || !/\/\s*>$/u.test(children[0].xml)) return false;
  return hasCanonicalStyleReferenceColor(children[0].xml);
}

function hasCanonicalStyleLineReference(rawXml) {
  const styles = directPresentationChildren(String(rawXml || ""), "cxnSp")
    .filter((child) => child.localName.toLowerCase() === "style");
  if (styles.length !== 1) return false;
  const styleOpen = /^<[^>]+>/u.exec(styles[0].xml)?.[0] || "";
  if (!styleOpen || !hasAllowedNativeNamespaceAttributes(styleOpen) || nativeStyleAttributes(styleOpen).length !== 0) return false;
  const references = directPresentationChildren(styles[0].xml, "style");
  if (references.length !== 4) return false;
  const expected = new Set(["lnref", "fillref", "effectref", "fontref"]);
  if (new Set(references.map((reference) => reference.localName.toLowerCase())).size !== 4 ||
      references.some((reference) => !expected.has(reference.localName.toLowerCase()))) return false;
  return references.every((reference) => hasCanonicalStyleReference(reference));
}

function nativeLineArrowLeaf(xml, endpointName, leafKind) {
  const nodes = directPresentationChildren(xml, "ln").filter((child) => child.localName === endpointName);
  if (nodes.length !== 1) return undefined;
  const node = nodes[0];
  const open = /^<[^>]+>/u.exec(node.xml)?.[0] || "";
  if (!/\/\s*>$/u.test(node.xml) || directPresentationChildren(node.xml, endpointName).length !== 0) return undefined;
  const attributes = nativeTagAttributes(open);
  const typeAttributes = attributes.filter((attribute) => attribute.name.split(":").pop()?.toLowerCase() === "type");
  if (typeAttributes.length !== 1 || attributes.some((attribute) => !["type", "w", "len"].includes(attribute.name.split(":").pop()?.toLowerCase())) ||
      ["w", "len"].some((name) => attributes.filter((attribute) => attribute.name.split(":").pop()?.toLowerCase() === name).length > 1) ||
      attributes.filter((attribute) => ["w", "len"].includes(attribute.name.split(":").pop()?.toLowerCase())).some((attribute) => !["sm", "med", "lg"].includes(attribute.value))) return undefined;
  const value = nativeLineArrowToken(typeAttributes[0].value);
  return value ? { leafKind, value } : undefined;
}

function deriveNativeLineLeaves(rawXml, nativeKind) {
  if (nativeKind !== "connector") return undefined;
  const source = String(rawXml || "");
  const spPrMatches = [...source.matchAll(NATIVE_SPPR_TAG)];
  if (spPrMatches.length !== 1) return undefined;
  const spPr = spPrMatches[0];
  const lineMatches = [...(spPr.groups?.value || "").matchAll(NATIVE_LINE_TAG)];
  if (lineMatches.length !== 1) return undefined;
  const line = lineMatches[0];
  const leaves = [];
  const width = nativeTagAttributes(line.groups?.attributes || "")
    .find((attribute) => attribute.name.split(":").pop()?.toLowerCase() === "w")?.value;
  if (width !== undefined && /^\d+$/u.test(width)) {
    const numericWidth = Number(width);
    if (Number.isSafeInteger(numericWidth) && numericWidth <= MAX_NATIVE_LINE_WIDTH_EMU) {
      leaves.push({ lineLeafIndex: leaves.length, leafKind: "lineWidthEmu", value: width, expectedHash: sha256(width) });
    }
  }
  const linePrefix = line.groups?.prefix || "";
  const solidMatches = [...(line.groups?.value || "").matchAll(NATIVE_SOLID_FILL_TAG)]
    .filter((match) => (match.groups?.prefix || "") === linePrefix);
  if (solidMatches.length === 1) {
    const colors = [...(solidMatches[0].groups?.value || "").matchAll(NATIVE_COLOR_OPEN_TAG)]
      .filter((match) => (match.groups?.prefix || "") === linePrefix);
    if (colors.length === 1) {
      const attributes = nativeTagAttributes(colors[0].groups?.attributes || "");
      if (attributes.length === 1 && attributes[0].name.split(":").pop()?.toLowerCase() === "val" && attributes[0].value) {
        const colorName = colors[0].groups?.name?.toLowerCase();
        if (colorName === "srgbclr" && /^[0-9a-f]{6}$/iu.test(attributes[0].value)) {
          const value = attributes[0].value.toUpperCase();
          leaves.push({ lineLeafIndex: leaves.length, leafKind: "lineRgb", value, expectedHash: sha256(value) });
        } else if (colorName === "schemeclr") {
          const value = nativeSchemeColorToken(attributes[0].value);
          if (value) leaves.push({ lineLeafIndex: leaves.length, leafKind: "lineScheme", value, expectedHash: sha256(value) });
        }
      }
    }
  }
  const dashMatches = [...(line.groups?.value || "").matchAll(NATIVE_PRESET_DASH_TAG)]
    .filter((match) => (match.groups?.prefix || "") === linePrefix);
  if (dashMatches.length === 1 && solidMatches.length === 1) {
    const attributes = nativeTagAttributes(dashMatches[0].groups?.attributes || "");
    if (attributes.length === 1 && attributes[0].name.split(":").pop()?.toLowerCase() === "val") {
      const value = nativeLineStyleToken(attributes[0].value);
      if (value) leaves.push({ lineLeafIndex: leaves.length, leafKind: "lineStyle", value, expectedHash: sha256(value) });
    }
  }
  const capAttributes = nativeTagAttributes(line.groups?.attributes || "")
    .filter((attribute) => attribute.name.split(":").pop()?.toLowerCase() === "cap");
  if (capAttributes.length === 1 && solidMatches.length === 1) {
    const value = nativeLineCapToken(capAttributes[0].value);
    if (value) leaves.push({ lineLeafIndex: leaves.length, leafKind: "lineCap", value, expectedHash: sha256(value) });
  }
  const joinNodes = directPresentationChildren(line[0], "ln")
    .filter((child) => child.localName === "round" || child.localName === "bevel" || child.localName === "miter");
  const hasSimpleLinePaint = leaves.some((leaf) => leaf.leafKind === "lineRgb" || leaf.leafKind === "lineScheme");
  const hasStyleLineReference = hasCanonicalStyleLineReference(source);
  if (joinNodes.length === 1 && hasSimpleLinePaint) {
    const open = /^<[^>]+>/u.exec(joinNodes[0].xml)?.[0] || "";
    if (/\/\s*>$/u.test(joinNodes[0].xml) && nativeTagAttributes(open).length === 0 && directPresentationChildren(joinNodes[0].xml, joinNodes[0].localName).length === 0) {
      const value = nativeLineJoinToken(joinNodes[0].localName);
      if (value) leaves.push({ lineLeafIndex: leaves.length, leafKind: "lineJoin", value, expectedHash: sha256(value) });
    }
  }
  if (hasSimpleLinePaint || hasStyleLineReference) {
    for (const [endpointName, leafKind] of [["headEnd", "lineStartArrow"], ["tailEnd", "lineEndArrow"]]) {
      const leaf = nativeLineArrowLeaf(line[0], endpointName, leafKind);
      if (leaf) leaves.push({ lineLeafIndex: leaves.length, ...leaf, expectedHash: sha256(leaf.value) });
    }
  }
  return leaves.length ? Object.freeze(leaves) : undefined;
}

function nativeLineRecord(leaves) {
  return leaves ? Object.freeze(leaves.map((leaf) => Object.freeze({
    lineLeafIndex: leaf.lineLeafIndex,
    leafKind: leaf.leafKind || "lineRgb",
    value: leaf.leafKind === "lineWidthEmu"
      ? Number(leaf.value)
      : leaf.leafKind === "lineScheme" || leaf.leafKind === "lineStyle" || leaf.leafKind === "lineCap" || leaf.leafKind === "lineJoin" || leaf.leafKind === "lineStartArrow" || leaf.leafKind === "lineEndArrow" ? leaf.value : `#${leaf.value.toLowerCase()}`,
    expectedValue: leaf.value,
    expectedHash: sha256(leaf.value),
  }))) : undefined;
}

function nativeLineEditableFields(leaves) {
  return leaves?.length ? [...new Set(leaves.map((leaf) => leaf.leafKind || "lineRgb"))] : [];
}

// A group that the semantic importer cannot model can still expose a tiny,
// source-bound style surface. Walk only direct PresentationML children and
// issue leaves for unambiguous solid fills, outline colors, and canonical
// outline widths on descendant
// p:sp nodes. This keeps the group topology, effects, and every unsupported
// paint opaque while making common theme-driven template shapes reusable.
function deriveNativeStyleLeaves(rawXml, nativeKind) {
  if (nativeKind !== "group") return undefined;
  const fillLeaves = [];
  const lineLeaves = [];
  const lineWidthLeaves = [];
  const lineStyleLeaves = [];
  const lineCapLeaves = [];
  const lineJoinLeaves = [];
  const lineStartArrowLeaves = [];
  const lineEndArrowLeaves = [];
  const colorLeaf = (solid, prefix) => {
    const colors = directPresentationChildren(solid.xml, "solidFill")
      .filter((child) => child.localName === "schemeClr" || child.localName === "srgbClr");
    if (colors.length !== 1) return undefined;
    const color = colors[0];
    if (directPresentationChildren(color.xml, color.localName).length !== 0) return undefined;
    const open = /^<[^>]+>/u.exec(color.xml)?.[0];
    const attributes = nativeTagAttributes(open || "");
    if (attributes.length !== 1 || attributes[0].name.split(":").pop()?.toLowerCase() !== "val" || !attributes[0].value) return undefined;
    if (color.localName === "schemeClr") {
      const value = nativeSchemeColorToken(attributes[0].value);
      return value ? { leafKind: `${prefix}Scheme`, value } : undefined;
    }
    if (!/^[0-9a-f]{6}$/iu.test(attributes[0].value)) return undefined;
    return { leafKind: `${prefix}Rgb`, value: attributes[0].value.toUpperCase() };
  };
  const visitGroup = (xml) => {
    for (const child of directPresentationChildren(xml, "grpSp")) {
      if (child.localName === "grpSp") {
        visitGroup(child.xml);
        continue;
      }
      if (child.localName !== "sp") continue;
      const shapeProperties = directPresentationChildren(child.xml, "sp").find((entry) => entry.localName === "spPr");
      if (!shapeProperties) continue;
      const fillNodes = directPresentationChildren(shapeProperties.xml, "spPr")
        .filter((entry) => ["noFill", "solidFill", "gradFill", "blipFill", "pattFill"].includes(entry.localName));
      if (fillNodes.length === 1 && fillNodes[0].localName === "solidFill") {
        const leaf = colorLeaf(fillNodes[0], "fill");
        if (leaf) fillLeaves.push(leaf);
      }
      const outlines = directPresentationChildren(shapeProperties.xml, "spPr").filter((entry) => entry.localName === "ln");
      if (outlines.length !== 1) continue;
      const lineOpen = /^<[^>]+>/u.exec(outlines[0].xml)?.[0];
      const widthAttributes = nativeTagAttributes(lineOpen || "")
        .filter((attribute) => attribute.name.split(":").pop()?.toLowerCase() === "w");
      if (widthAttributes.length === 1 && /^[1-9]\d*$/u.test(widthAttributes[0].value)) {
        const width = Number(widthAttributes[0].value);
        if (Number.isSafeInteger(width) && width <= MAX_NATIVE_LINE_WIDTH_EMU) {
          lineWidthLeaves.push({ leafKind: "lineWidthEmu", value: widthAttributes[0].value });
        }
      }
      const lineFills = directPresentationChildren(outlines[0].xml, "ln")
        .filter((entry) => ["noFill", "solidFill", "gradFill", "blipFill", "pattFill"].includes(entry.localName));
      if (lineFills.length !== 1 || lineFills[0].localName !== "solidFill") continue;
      const lineLeaf = colorLeaf(lineFills[0], "line");
      if (lineLeaf) lineLeaves.push(lineLeaf);
      const dashNodes = directPresentationChildren(outlines[0].xml, "ln")
        .filter((entry) => entry.localName === "prstDash");
      if (dashNodes.length === 1) {
        const open = /^<[^>]+>/u.exec(dashNodes[0].xml)?.[0] || "";
        const attributes = nativeTagAttributes(open);
        if (attributes.length === 1 && attributes[0].name.split(":").pop()?.toLowerCase() === "val") {
          const value = nativeLineStyleToken(attributes[0].value);
          if (value) lineStyleLeaves.push({ leafKind: "lineStyle", value });
        }
      }
      const capAttributes = nativeTagAttributes(lineOpen || "")
        .filter((attribute) => attribute.name.split(":").pop()?.toLowerCase() === "cap");
      if (capAttributes.length === 1 && nativeLineCapToken(capAttributes[0].value) && lineLeaf) {
        lineCapLeaves.push({ leafKind: "lineCap", value: nativeLineCapToken(capAttributes[0].value) });
      }
      const joinNodes = directPresentationChildren(outlines[0].xml, "ln")
        .filter((entry) => entry.localName === "round" || entry.localName === "bevel" || entry.localName === "miter");
      if (joinNodes.length === 1 && lineLeaf) {
        const open = /^<[^>]+>/u.exec(joinNodes[0].xml)?.[0] || "";
        if (/\/\s*>$/u.test(joinNodes[0].xml) && nativeTagAttributes(open).length === 0 && directPresentationChildren(joinNodes[0].xml, joinNodes[0].localName).length === 0) {
          const value = nativeLineJoinToken(joinNodes[0].localName);
          if (value) lineJoinLeaves.push({ leafKind: "lineJoin", value });
        }
      }
      if (lineLeaf) {
        const startArrow = nativeLineArrowLeaf(outlines[0].xml, "headEnd", "lineStartArrow");
        if (startArrow) lineStartArrowLeaves.push(startArrow);
        const endArrow = nativeLineArrowLeaf(outlines[0].xml, "tailEnd", "lineEndArrow");
        if (endArrow) lineEndArrowLeaves.push(endArrow);
      }
    }
  };
  visitGroup(String(rawXml || ""));
  // Keep prior color indexes stable; append line widths as a separate family.
  const leaves = [...fillLeaves, ...lineLeaves, ...lineWidthLeaves, ...lineStyleLeaves, ...lineCapLeaves, ...lineJoinLeaves, ...lineStartArrowLeaves, ...lineEndArrowLeaves].map((leaf, nativeLeafIndex) => ({
    nativeLeafIndex,
    ...leaf,
    expectedHash: sha256(leaf.value),
  }));
  return leaves.length && leaves.length <= MAX_NATIVE_STYLE_LEAVES ? Object.freeze(leaves) : undefined;
}

function nativeStyleRecord(leaves) {
  return leaves ? Object.freeze(leaves.map((leaf) => Object.freeze({
    nativeLeafIndex: leaf.nativeLeafIndex,
    leafKind: leaf.leafKind,
    value: leaf.leafKind === "fillScheme" || leaf.leafKind === "lineScheme" || leaf.leafKind === "lineStyle" || leaf.leafKind === "lineCap" || leaf.leafKind === "lineJoin" || leaf.leafKind === "lineStartArrow" || leaf.leafKind === "lineEndArrow"
      ? leaf.value
      : leaf.leafKind === "lineWidthEmu" ? Number(leaf.value) : `#${leaf.value.toLowerCase()}`,
    expectedValue: leaf.value,
    expectedHash: sha256(leaf.value),
  }))) : undefined;
}

function nativeStyleEditableFields(leaves) {
  return leaves?.length ? [...new Set(leaves.map((leaf) => leaf.leafKind))] : [];
}

function normalizeNativeChart(config) {
  if (!config) return undefined;
  const partPath = String(config.partPath || "");
  const contentType = String(config.contentType || "").toLowerCase();
  const sourceSha256 = String(config.sourceSha256 || "").toLowerCase();
  const relationshipId = String(config.relationshipId || "");
  const sourceLeaves = config.titleLeaves;
  if (!partPath || contentType !== CHART_CONTENT_TYPE || !/^[0-9a-f]{64}$/iu.test(sourceSha256) ||
      !relationshipId || !Array.isArray(sourceLeaves) || sourceLeaves.length > 256) {
    throw new TypeError("Native chart title binding is incomplete or outside the bounded profile.");
  }
  const titleLeaves = sourceLeaves.map((leaf, index) => {
    const textLeafIndex = Number(leaf?.textLeafIndex);
    const text = String(leaf?.text ?? "");
    if (textLeafIndex !== index || text.length > 32_767 || !validDiagramNodeText(text)) {
      throw new TypeError("Native chart title binding contains an invalid text leaf.");
    }
    return Object.freeze({ textLeafIndex, text });
  });
  const embeddedPackagePartPath = String(config.embeddedPackagePartPath || "");
  const embeddedPackageSourceSha256 = String(config.embeddedPackageSourceSha256 || "").toLowerCase();
  const embeddedPackageRelationshipId = String(config.embeddedPackageRelationshipId || "");
  const sourceDataPoints = config.dataPoints || [];
  if (!Array.isArray(sourceDataPoints) || sourceDataPoints.length > 256 ||
      (sourceDataPoints.length > 0 && (!embeddedPackagePartPath || !/^[0-9a-f]{64}$/u.test(embeddedPackageSourceSha256) || !embeddedPackageRelationshipId))) {
    throw new TypeError("Native chart data binding is incomplete or outside the bounded profile.");
  }
  const dataPoints = sourceDataPoints.map((point) => {
    const seriesIndex = Number(point?.seriesIndex);
    const pointIndex = Number(point?.pointIndex);
    const value = String(point?.value ?? "");
    const formula = String(point?.formula || "");
    const worksheetPartPath = String(point?.worksheetPartPath || "");
    const worksheetSourceSha256 = String(point?.worksheetSourceSha256 || "").toLowerCase();
    const worksheetName = String(point?.worksheetName || "");
    const cellReference = String(point?.cellReference || "");
    if (!Number.isSafeInteger(seriesIndex) || seriesIndex < 0 || !Number.isSafeInteger(pointIndex) || pointIndex < 0 ||
        !validNativeChartNumber(value) || !formula || !worksheetPartPath || !/^[0-9a-f]{64}$/u.test(worksheetSourceSha256) ||
        !worksheetName || !/^[A-Z]{1,3}[1-9][0-9]*$/u.test(cellReference)) {
      throw new TypeError("Native chart data binding contains an invalid point.");
    }
    return Object.freeze({ seriesIndex, pointIndex, value, formula, worksheetPartPath, worksheetSourceSha256, worksheetName, cellReference });
  });
  if (new Set(dataPoints.map((point) => `${point.seriesIndex}:${point.pointIndex}`)).size !== dataPoints.length) {
    throw new TypeError("Native chart data binding contains duplicate points.");
  }
  if (titleLeaves.length === 0 && dataPoints.length === 0) {
    throw new TypeError("Native chart binding contains no safe leaf.");
  }
  return Object.freeze({
    partPath,
    contentType,
    sourceSha256,
    relationshipId,
    titleLeaves: Object.freeze(titleLeaves),
    embeddedPackagePartPath,
    embeddedPackageSourceSha256,
    embeddedPackageRelationshipId,
    dataPoints: Object.freeze(dataPoints),
  });
}

function validNativeChartNumber(value) {
  return value.length > 0 && value.length <= 128 && /^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[Ee][+-]?[0-9]+)?$/u.test(value) && Number.isFinite(Number(value));
}

function nativeChartRecord(binding, leaves = binding?.titleLeaves, dataPoints = binding?.dataPoints) {
  if (!binding) return undefined;
  return Object.freeze({
    partPath: binding.partPath,
    contentType: binding.contentType,
    sourceSha256: binding.sourceSha256,
    relationshipId: binding.relationshipId,
    titleLeaves: Object.freeze(leaves.map((leaf) => Object.freeze({
      textLeafIndex: leaf.textLeafIndex,
      text: leaf.text,
    }))),
    embeddedPackagePartPath: binding.embeddedPackagePartPath,
    embeddedPackageSourceSha256: binding.embeddedPackageSourceSha256,
    embeddedPackageRelationshipId: binding.embeddedPackageRelationshipId,
    dataPoints: Object.freeze(dataPoints.map((point) => Object.freeze({ ...point }))),
  });
}

function normalizeOleOfficePackage(config) {
  if (!config) return undefined;
  const partPath = String(config.partPath || "");
  const contentType = String(config.contentType || "").toLowerCase();
  const sourceSha256 = String(config.sourceSha256 || "").toLowerCase();
  const relationshipId = String(config.relationshipId || "");
  const kind = String(config.kind || "").toLowerCase();
  if (!partPath || contentType !== DOCX_CONTENT_TYPE || !/^[0-9a-f]{64}$/i.test(sourceSha256) || !relationshipId || kind !== "docx") {
    throw new TypeError("Embedded Office package binding is incomplete or outside the bounded DOCX profile.");
  }
  return Object.freeze({ partPath, contentType, sourceSha256, relationshipId, kind });
}

function hasOnlyValidUnicodeScalars(value) {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      index += 1;
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      return false;
    }
  }
  return true;
}

function validDiagramNodeText(value) {
  return value.length <= MAX_DIAGRAM_NODE_TEXT_LENGTH &&
    !/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/u.test(value) &&
    hasOnlyValidUnicodeScalars(value);
}

function validDiagramModelId(value) {
  if (value.length > 1_024 || /[\u0000-\u001f]/u.test(value) || !hasOnlyValidUnicodeScalars(value)) return false;
  if (/^[+-]?\d+$/u.test(value)) {
    const numeric = BigInt(value);
    return numeric >= -2_147_483_648n && numeric <= 2_147_483_647n;
  }
  return /^\{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}$/iu.test(value);
}

function normalizeDiagramText(config) {
  if (!config) return undefined;
  const partPath = String(config.partPath || "");
  const contentType = String(config.contentType || "");
  const sourceSha256 = String(config.sourceSha256 || "").toLowerCase();
  const relationshipId = String(config.relationshipId || "");
  const nodes = config.nodes;
  if (!partPath || !contentType || !/^[0-9a-f]{64}$/i.test(sourceSha256) || !relationshipId || !Array.isArray(nodes) || !nodes.length) {
    throw new TypeError("SmartArt diagram text binding is incomplete.");
  }
  const seen = new Set();
  const normalizedNodes = nodes.map((node) => {
    const id = String(node?.id ?? node?.modelId ?? "");
    const text = String(node?.text ?? "");
    const sourceRuns = node?.runs ?? node?.runTexts;
    const runs = Array.isArray(sourceRuns) && sourceRuns.length
      ? sourceRuns.map((value) => String(value ?? ""))
      : [text];
    if (!id || !validDiagramModelId(id) || !validDiagramNodeText(text) ||
        runs.length > MAX_DIAGRAM_NODE_RUNS || runs.some((value) => !validDiagramNodeText(value)) ||
        runs.join("") !== text || seen.has(id)) {
      throw new TypeError("SmartArt diagram text binding contains an invalid node.");
    }
    seen.add(id);
    return Object.freeze({ id, text, runs: Object.freeze(runs) });
  });
  return Object.freeze({
    partPath,
    contentType,
    sourceSha256,
    relationshipId,
    nodes: Object.freeze(normalizedNodes),
  });
}

function diagramTextRecord(binding, nodes) {
  if (!binding) return undefined;
  return Object.freeze({
    partPath: binding.partPath,
    contentType: binding.contentType,
    sourceSha256: binding.sourceSha256,
    relationshipId: binding.relationshipId,
    nodes: Object.freeze(nodes.map((node) => Object.freeze({
      id: node.id,
      text: node.text,
      runs: Object.freeze(node.runs.map((text) => text)),
    }))),
  });
}

export function createNativePresentationObjectClass({ normalizeFrame }) {
  return class NativePresentationObject {
    constructor(slide, config = {}) {
      this.slide = slide;
      this.kind = "nativeObject";
      this.id = config.id || aid("no");
      this.nativeId = config.nativeId;
      this.creationId = config.creationId;
      this.name = config.name || "";
      this.nativeKind = config.nativeKind || "graphicFrame";
      this.position = normalizeFrame(config, { left: 0, top: 0, width: 1, height: 1 });
      this.rawXml = String(config.rawXml || "");
      const nativeText = String(config.text ?? "");
      Object.defineProperty(this, "text", {
        configurable: false,
        enumerable: true,
        writable: false,
        value: nativeText.slice(0, MAX_NATIVE_TEXT_LENGTH),
      });
      Object.defineProperty(this, "textTruncated", {
        configurable: false,
        enumerable: true,
        writable: false,
        value: nativeText.length > MAX_NATIVE_TEXT_LENGTH,
      });
      Object.defineProperty(this, "textLength", {
        configurable: false,
        enumerable: true,
        writable: false,
        value: nativeText.length,
      });
      this.sourcePart = config.sourcePart;
      Object.defineProperty(this, "editable", { enumerable: true, value: false, writable: false });
      const placement = config.placementCapability || {};
      const placementCapability = Object.freeze({
        sourceBound: placement.sourceBound === true,
        known: placement.known !== false,
        supported: placement.supported === true,
        blockedReason: String(placement.blockedReason || ""),
        sourceRevisionSha256: placement.sourceRevisionSha256 ? String(placement.sourceRevisionSha256).toLowerCase() : undefined,
      });
      Object.defineProperty(this, "placementCapability", {
        configurable: false,
        enumerable: true,
        writable: false,
        value: placementCapability,
      });
      Object.defineProperty(this, "_nativePlacementSourcePosition", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: Object.freeze({ ...this.position }),
      });
      Object.defineProperty(this, "_nativePlacementMutationIssued", {
        configurable: false,
        enumerable: false,
        writable: true,
        value: false,
      });
      this.relationshipReferences = (config.relationshipReferences || []).map((reference) => ({ ...reference }));
      this.rootRelationships = (config.rootRelationships || []).map((relationship) => ({ ...relationship }));
      const shareImportedPartBytes = config._officeKitSharePartBytes === true;
      this.parts = (config.parts || []).map((part) => ({
        ...part,
        bytes: shareImportedPartBytes ? part.bytes : new Uint8Array(part.bytes),
        relationships: (part.relationships || []).map((relationship) => ({ ...relationship })),
      }));
      const oleWorkbook = config.oleWorkbook ? Object.freeze({
        partPath: String(config.oleWorkbook.partPath || ""),
        contentType: String(config.oleWorkbook.contentType || ""),
        sourceSha256: String(config.oleWorkbook.sourceSha256 || "").toLowerCase(),
        relationshipId: String(config.oleWorkbook.relationshipId || ""),
      }) : undefined;
      Object.defineProperty(this, "oleWorkbook", {
        configurable: false,
        enumerable: true,
        writable: false,
        value: oleWorkbook,
      });
      const oleOfficePackage = normalizeOleOfficePackage(config.oleOfficePackage);
      Object.defineProperty(this, "oleOfficePackage", {
        configurable: false,
        enumerable: true,
        writable: false,
        value: oleOfficePackage,
      });
      const diagramText = normalizeDiagramText(config.diagramText);
      Object.defineProperty(this, "_diagramTextBinding", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: diagramText,
      });
      Object.defineProperty(this, "_diagramTextNodes", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: diagramText ? diagramText.nodes.map((node) => ({ ...node, runs: [...node.runs] })) : undefined,
      });
      const nativeChart = normalizeNativeChart(config.nativeChart);
      Object.defineProperty(this, "_nativeChartBinding", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeChart,
      });
      Object.defineProperty(this, "_nativeChartTitleLeaves", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeChart ? nativeChart.titleLeaves.map((leaf) => ({ ...leaf })) : undefined,
      });
      Object.defineProperty(this, "_nativeChartDataPoints", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeChart ? nativeChart.dataPoints.map((point) => ({ ...point })) : undefined,
      });
      const nativeTextBinding = deriveNativeTextLeaves(this.rawXml, this.nativeKind);
      Object.defineProperty(this, "_nativeTextBinding", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeTextBinding,
      });
      Object.defineProperty(this, "_nativeTextLeaves", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeTextBinding ? nativeTextBinding.map((leaf) => ({ ...leaf })) : undefined,
      });
      const nativeLineBinding = deriveNativeLineLeaves(this.rawXml, this.nativeKind);
      Object.defineProperty(this, "_nativeLineBinding", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeLineBinding,
      });
      Object.defineProperty(this, "_nativeLineLeaves", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeLineBinding ? nativeLineBinding.map((leaf) => ({ ...leaf })) : undefined,
      });
      const nativeStyleBinding = deriveNativeStyleLeaves(this.rawXml, this.nativeKind);
      Object.defineProperty(this, "_nativeStyleBinding", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeStyleBinding,
      });
      Object.defineProperty(this, "_nativeStyleLeaves", {
        configurable: false,
        enumerable: false,
        writable: false,
        value: nativeStyleBinding ? nativeStyleBinding.map((leaf) => ({ ...leaf })) : undefined,
      });
      Object.defineProperty(this, "diagramText", {
        configurable: false,
        enumerable: true,
        get: () => diagramTextRecord(this._diagramTextBinding, this._diagramTextNodes || []),
      });
      Object.defineProperty(this, "nativeTextLeaves", {
        configurable: false,
        enumerable: true,
        get: () => nativeTextRecord(this._nativeTextLeaves),
      });
      Object.defineProperty(this, "nativeLineLeaves", {
        configurable: false,
        enumerable: true,
        get: () => nativeLineRecord(this._nativeLineLeaves),
      });
      Object.defineProperty(this, "nativeStyleLeaves", {
        configurable: false,
        enumerable: true,
        get: () => nativeStyleRecord(this._nativeStyleLeaves),
      });
      Object.defineProperty(this, "_embeddedWorkbookReplacement", {
        configurable: false,
        enumerable: false,
        writable: true,
        value: undefined,
      });
      Object.defineProperty(this, "_embeddedOfficePackageReplacement", {
        configurable: false,
        enumerable: false,
        writable: true,
        value: undefined,
      });
    }

    setName(value) {
      if (!this.editable) throw new Error(`Native ${this.nativeKind} object ${this.id} is read-only.`);
      const name = String(value ?? "");
      if (name.length > 1_024) throw new RangeError("Native presentation object names cannot exceed 1024 characters.");
      this.name = name;
      return this;
    }

    setPosition(value = {}) {
      if (!this.editable && !this.placementCapability.supported) {
        const reason = this.placementCapability.blockedReason ? `: ${this.placementCapability.blockedReason}` : ".";
        throw new Error(`Native ${this.nativeKind} object ${this.id} is read-only for placement${reason}`);
      }
      this.position = normalizeFrame({ position: { ...this.position, ...value } }, this.position);
      this._nativePlacementMutationIssued = true;
      return this;
    }

    _nativePlacementChanged() {
      const source = this._nativePlacementSourcePosition;
      return this.position.left !== source.left || this.position.top !== source.top ||
        this.position.width !== source.width || this.position.height !== source.height;
    }

    embeddedWorkbookPart() {
      if (!this.oleWorkbook) throw new Error(`Native ${this.nativeKind} object ${this.id} has no embedded XLSX workbook.`);
      const matches = this.parts.filter((part) => part.path === this.oleWorkbook.partPath && part.contentType === this.oleWorkbook.contentType);
      if (matches.length !== 1) throw new Error(`Native ${this.nativeKind} object ${this.id} no longer resolves to one embedded XLSX workbook part.`);
      return matches[0];
    }

    getEmbeddedWorkbook() {
      const part = this.embeddedWorkbookPart();
      const replacement = this._embeddedWorkbookReplacement;
      return new FileBlob(Uint8Array.from(replacement || part.bytes), {
        type: this.oleWorkbook.contentType,
        metadata: replacement
          ? { artifactKind: "workbook", source: "presentationOleObject", partPath: this.oleWorkbook.partPath, boundSourceSha256: this.oleWorkbook.sourceSha256, pendingReplacement: true }
          : { artifactKind: "workbook", source: "presentationOleObject", partPath: this.oleWorkbook.partPath, sourceSha256: this.oleWorkbook.sourceSha256 },
      });
    }

    replaceEmbeddedWorkbook(input) {
      this.embeddedWorkbookPart();
      if (input == null || typeof input === "string" || !(input instanceof FileBlob || input instanceof ArrayBuffer || input instanceof Uint8Array || ArrayBuffer.isView(input))) {
        throw new TypeError("Embedded workbook replacement must be a FileBlob, Uint8Array, ArrayBuffer, or ArrayBuffer view.");
      }
      const bytes = input instanceof FileBlob ? input.bytes : toUint8Array(input);
      if (!bytes.byteLength || bytes.byteLength > MAX_EMBEDDED_WORKBOOK_BYTES) {
        throw new RangeError(`Embedded workbook replacement must contain 1 through ${MAX_EMBEDDED_WORKBOOK_BYTES} bytes.`);
      }
      this._embeddedWorkbookReplacement = Uint8Array.from(bytes);
      return this;
    }

    _embeddedWorkbookReplacementBytes() {
      return this._embeddedWorkbookReplacement ? Uint8Array.from(this._embeddedWorkbookReplacement) : undefined;
    }

    embeddedOfficePackagePart() {
      if (!this.oleOfficePackage) throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded embedded Office package.`);
      const matches = this.parts.filter((part) => part.path === this.oleOfficePackage.partPath && part.contentType === this.oleOfficePackage.contentType);
      if (matches.length !== 1) throw new Error(`Native ${this.nativeKind} object ${this.id} no longer resolves to one embedded Office package part.`);
      return matches[0];
    }

    getEmbeddedOfficePackage() {
      if (this.oleWorkbook) {
        const workbook = this.getEmbeddedWorkbook();
        return new FileBlob(workbook.bytes, {
          type: workbook.type,
          metadata: { ...workbook.metadata, artifactKind: "officePackage", officePackageKind: "xlsx" },
        });
      }
      const part = this.embeddedOfficePackagePart();
      const replacement = this._embeddedOfficePackageReplacement;
      return new FileBlob(Uint8Array.from(replacement || part.bytes), {
        type: this.oleOfficePackage.contentType,
        metadata: replacement
          ? { artifactKind: "officePackage", officePackageKind: this.oleOfficePackage.kind, source: "presentationOleObject", partPath: this.oleOfficePackage.partPath, boundSourceSha256: this.oleOfficePackage.sourceSha256, pendingReplacement: true }
          : { artifactKind: "officePackage", officePackageKind: this.oleOfficePackage.kind, source: "presentationOleObject", partPath: this.oleOfficePackage.partPath, sourceSha256: this.oleOfficePackage.sourceSha256 },
      });
    }

    replaceEmbeddedOfficePackage(input) {
      if (this.oleWorkbook) return this.replaceEmbeddedWorkbook(input);
      this.embeddedOfficePackagePart();
      if (input == null || typeof input === "string" || !(input instanceof FileBlob || input instanceof ArrayBuffer || input instanceof Uint8Array || ArrayBuffer.isView(input))) {
        throw new TypeError("Embedded Office package replacement must be a FileBlob, Uint8Array, ArrayBuffer, or ArrayBuffer view.");
      }
      if (input instanceof FileBlob && String(input.type || "").toLowerCase() !== this.oleOfficePackage.contentType) {
        throw new TypeError(`Embedded Office package replacement must retain content type ${this.oleOfficePackage.contentType}.`);
      }
      const bytes = input instanceof FileBlob ? input.bytes : toUint8Array(input);
      if (!bytes.byteLength || bytes.byteLength > MAX_EMBEDDED_OFFICE_PACKAGE_BYTES) {
        throw new RangeError(`Embedded Office package replacement must contain 1 through ${MAX_EMBEDDED_OFFICE_PACKAGE_BYTES} bytes.`);
      }
      this._embeddedOfficePackageReplacement = Uint8Array.from(bytes);
      return this;
    }

    _embeddedOfficePackageReplacementBytes() {
      return this._embeddedOfficePackageReplacement ? Uint8Array.from(this._embeddedOfficePackageReplacement) : undefined;
    }

    setDiagramNodeText(nodeId, value) {
      if (!this._diagramTextBinding || !this._diagramTextNodes) {
        throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded SmartArt diagram-text capability.`);
      }
      const id = String(nodeId ?? "");
      const text = String(value ?? "");
      if (!validDiagramNodeText(text)) {
        throw new RangeError(`SmartArt node text must contain at most ${MAX_DIAGRAM_NODE_TEXT_LENGTH} XML-safe characters.`);
      }
      const node = this._diagramTextNodes.find((candidate) => candidate.id === id);
      if (!node) throw new Error(`SmartArt node ${id || "(empty)"} is not part of the source-bound diagram profile.`);
      if (node.runs.length !== 1) {
        throw new Error(`SmartArt node ${id} has ${node.runs.length} source-bound styled runs; use setDiagramNodeRunText() so OfficeKit does not guess a formatting boundary.`);
      }
      node.text = text;
      node.runs[0] = text;
      return this;
    }

    setDiagramNodeRunText(nodeId, runIndex, value) {
      if (!this._diagramTextBinding || !this._diagramTextNodes) {
        throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded SmartArt diagram-text capability.`);
      }
      const id = String(nodeId ?? "");
      const index = Number(runIndex);
      const text = String(value ?? "");
      if (!Number.isSafeInteger(index) || index < 0) throw new TypeError("SmartArt runIndex must be a non-negative integer.");
      if (!validDiagramNodeText(text)) {
        throw new RangeError(`SmartArt run text must contain at most ${MAX_DIAGRAM_NODE_TEXT_LENGTH} XML-safe characters.`);
      }
      const node = this._diagramTextNodes.find((candidate) => candidate.id === id);
      if (!node) throw new Error(`SmartArt node ${id || "(empty)"} is not part of the source-bound diagram profile.`);
      if (index >= node.runs.length) throw new RangeError(`SmartArt node ${id} has no source-bound run at index ${index}.`);
      const runs = node.runs.map((current, candidate) => candidate === index ? text : current);
      const combined = runs.join("");
      if (!validDiagramNodeText(combined)) {
        throw new RangeError(`SmartArt node text must contain at most ${MAX_DIAGRAM_NODE_TEXT_LENGTH} XML-safe characters across all runs.`);
      }
      node.runs[index] = text;
      node.text = combined;
      return this;
    }

    _diagramTextSourceBinding() {
      return this._diagramTextBinding ? diagramTextRecord(this._diagramTextBinding, this._diagramTextBinding.nodes) : undefined;
    }

    _diagramTextReplacement() {
      if (!this._diagramTextBinding || !this._diagramTextNodes) return undefined;
      const changed = this._diagramTextNodes.some((node, index) =>
        node.text !== this._diagramTextBinding.nodes[index].text ||
        node.runs.some((text, runIndex) => text !== this._diagramTextBinding.nodes[index].runs[runIndex]));
      return changed ? diagramTextRecord(this._diagramTextBinding, this._diagramTextNodes) : undefined;
    }

    _diagramTextRunRecords() {
      if (!this._diagramTextBinding || !this._diagramTextNodes) return undefined;
      let textLeafIndex = 0;
      return Object.freeze(this._diagramTextNodes.flatMap((node, nodeIndex) =>
        node.runs.map((text, runIndex) => Object.freeze({
          textLeafIndex: textLeafIndex++,
          nodeId: node.id,
          nodeIndex,
          runIndex,
          text,
        }))));
    }

    _setDiagramTextRun(nodeId, runIndex, value) {
      this.setDiagramNodeRunText(nodeId, runIndex, value);
    }

    _nativeChartSourceBinding() {
      return nativeChartRecord(this._nativeChartBinding);
    }

    _nativeChartTitleRecords() {
      return this._nativeChartBinding ? nativeChartRecord(this._nativeChartBinding, this._nativeChartTitleLeaves).titleLeaves : undefined;
    }

    _nativeChartDataPointRecords() {
      return this._nativeChartBinding ? Object.freeze(this._nativeChartDataPoints.map((point) => Object.freeze({ ...point }))) : undefined;
    }

    _nativeChartCurrentRecord() {
      return this._nativeChartBinding ? nativeChartRecord(this._nativeChartBinding, this._nativeChartTitleLeaves, this._nativeChartDataPoints) : undefined;
    }

    _nativeTextSourceBinding() {
      return nativeTextRecord(this._nativeTextBinding);
    }

    _nativeTextRecords() {
      return nativeTextRecord(this._nativeTextLeaves);
    }

    _nativeLineSourceBinding() {
      return nativeLineRecord(this._nativeLineBinding);
    }

    _nativeLineRecords() {
      return nativeLineRecord(this._nativeLineLeaves);
    }

    _nativeStyleSourceBinding() {
      return nativeStyleRecord(this._nativeStyleBinding);
    }

    _nativeStyleRecords() {
      return nativeStyleRecord(this._nativeStyleLeaves);
    }

    _setNativeLineLeaf(index, value) {
      if (!this._nativeLineBinding || !this._nativeLineLeaves || !this._nativeLineLeaves[index]) {
        throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded native line leaf ${index}.`);
      }
      const color = String(value ?? "").trim();
      const leafKind = this._nativeLineBinding[index].leafKind || "lineRgb";
      if (leafKind === "lineStyle") {
        const style = nativeLineStyleToken(color);
        if (!style || style !== color) throw new RangeError("Native line style requires solid, dashed, dotted, dash-dot, or dash-dot-dot.");
        this._nativeLineLeaves[index].value = style;
        return;
      }
      if (leafKind === "lineCap") {
        const cap = nativeLineCapToken(color);
        if (!cap || cap !== color) throw new RangeError("Native line cap requires flat, round, or square.");
        this._nativeLineLeaves[index].value = cap;
        return;
      }
      if (leafKind === "lineJoin") {
        const join = nativeLineJoinToken(color);
        if (!join || join !== color) throw new RangeError("Native line join requires round, bevel, or miter.");
        this._nativeLineLeaves[index].value = join;
        return;
      }
      if (leafKind === "lineStartArrow" || leafKind === "lineEndArrow") {
        const arrow = nativeLineArrowToken(color);
        if (!arrow || arrow !== color) throw new RangeError("Native line arrow requires none, triangle, stealth, diamond, oval, or arrow.");
        this._nativeLineLeaves[index].value = arrow;
        return;
      }
      if (leafKind === "lineWidthEmu") {
        if (!/^\d+$/u.test(color)) throw new RangeError("Native line width requires a non-negative integer EMU value.");
        const width = Number(color);
        if (!Number.isSafeInteger(width) || width > MAX_NATIVE_LINE_WIDTH_EMU) {
          throw new RangeError("Native line width is outside the safe EMU range.");
        }
        this._nativeLineLeaves[index].value = color;
        return;
      }
      if (leafKind === "lineScheme") {
        const token = nativeSchemeColorToken(color);
        if (!token) throw new RangeError("Native line scheme color must be a supported theme token.");
        this._nativeLineLeaves[index].value = token;
        return;
      }
      if (!/^[0-9a-f]{6}$/iu.test(color)) {
        throw new RangeError("Native line color must be a six-digit RGB value.");
      }
      this._nativeLineLeaves[index].value = color.toUpperCase();
    }

    _setNativeStyleLeaf(index, value) {
      if (!this._nativeStyleBinding || !this._nativeStyleLeaves || !this._nativeStyleLeaves[index]) {
        throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded native style leaf ${index}.`);
      }
      const leafKind = this._nativeStyleBinding[index].leafKind;
      if (leafKind === "lineStyle") {
        const style = nativeLineStyleToken(String(value ?? "").trim());
        if (!style || style !== String(value ?? "").trim()) throw new RangeError("Native style line style requires solid, dashed, dotted, dash-dot, or dash-dot-dot.");
        this._nativeStyleLeaves[index].value = style;
        return;
      }
      if (leafKind === "lineCap") {
        const token = String(value ?? "").trim();
        const cap = nativeLineCapToken(token);
        if (!cap || cap !== token) throw new RangeError("Native style line cap requires flat, round, or square.");
        this._nativeStyleLeaves[index].value = cap;
        return;
      }
      if (leafKind === "lineJoin") {
        const token = String(value ?? "").trim();
        const join = nativeLineJoinToken(token);
        if (!join || join !== token) throw new RangeError("Native style line join requires round, bevel, or miter.");
        this._nativeStyleLeaves[index].value = join;
        return;
      }
      if (leafKind === "lineStartArrow" || leafKind === "lineEndArrow") {
        const token = String(value ?? "").trim();
        const arrow = nativeLineArrowToken(token);
        if (!arrow || arrow !== token) throw new RangeError("Native style line arrow requires none, triangle, stealth, diamond, oval, or arrow.");
        this._nativeStyleLeaves[index].value = arrow;
        return;
      }
      if (leafKind === "lineWidthEmu") {
        const token = String(value ?? "").trim();
        if (!/^(?:0|[1-9]\d*)$/u.test(token)) throw new RangeError("Native style line width requires a non-negative integer EMU value.");
        const width = Number(token);
        if (!Number.isSafeInteger(width) || width > MAX_NATIVE_LINE_WIDTH_EMU) {
          throw new RangeError("Native style line width is outside the safe EMU range.");
        }
        this._nativeStyleLeaves[index].value = token;
        return;
      }
      if (leafKind === "fillScheme" || leafKind === "lineScheme") {
        const token = nativeSchemeColorToken(String(value ?? "").trim());
        if (!token) throw new RangeError("Native style scheme color must be a supported theme token.");
        this._nativeStyleLeaves[index].value = token;
        return;
      }
      const color = String(value ?? "").trim().replace(/^#/u, "");
      if (!/^[0-9a-f]{6}$/iu.test(color)) throw new RangeError("Native style color must be a six-digit RGB value.");
      this._nativeStyleLeaves[index].value = color.toUpperCase();
    }

    _setNativeTextLeaf(index, value) {
      if (!this._nativeTextBinding || !this._nativeTextLeaves || !this._nativeTextLeaves[index]) {
        throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded native text leaf ${index}.`);
      }
      const text = String(value ?? "");
      if (!validDiagramNodeText(text)) {
        throw new RangeError(`Native text leaf must contain at most ${MAX_NATIVE_TEXT_LENGTH} XML-safe characters.`);
      }
      this._nativeTextLeaves[index].text = text;
    }

    get deletionCapability() {
      return presentationElementDeletionCapability(this, "native object");
    }

    _setNativeChartTitleLeaf(index, value) {
      if (!this._nativeChartBinding || !this._nativeChartTitleLeaves || !this._nativeChartTitleLeaves[index]) {
        throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded chart-title leaf ${index}.`);
      }
      this._nativeChartTitleLeaves[index].text = value;
    }

    _setNativeChartDataPoint(seriesIndex, pointIndex, value) {
      const point = this._nativeChartDataPoints?.find((candidate) => candidate.seriesIndex === seriesIndex && candidate.pointIndex === pointIndex);
      if (!this._nativeChartBinding || !point) {
        throw new Error(`Native ${this.nativeKind} object ${this.id} has no bounded chart data point ${seriesIndex}:${pointIndex}.`);
      }
      point.value = value;
    }

    inspectRecord() {
      const frame = this.parentGroup ? this.parentGroup.absoluteChildFrame(this) : this.position;
      const editableFields = [
        ...(this.oleWorkbook ? ["embeddedWorkbook"] : []),
        ...(this.oleOfficePackage ? ["embeddedOfficePackage"] : []),
        ...(this._diagramTextBinding ? ["diagramText"] : []),
        ...(this._nativeChartTitleLeaves?.length ? ["chartTitleText"] : []),
        ...(this._nativeChartDataPoints?.length ? ["chartDataValue"] : []),
        ...(this._nativeTextLeaves?.length ? ["nativeText"] : []),
        ...nativeLineEditableFields(this._nativeLineLeaves),
        ...nativeStyleEditableFields(this._nativeStyleLeaves),
        ...(this.placementCapability.supported ? ["position"] : []),
      ];
      return {
        kind: "nativeObject",
        id: this.id,
        slide: this.slide.index + 1,
        name: this.name || undefined,
        nativeKind: this.nativeKind,
        nativeId: this.nativeId,
        creationId: this.creationId,
        sourcePart: this.sourcePart,
        relationships: this.rootRelationships.length,
        preservedParts: this.parts.length,
        relationshipReferences: this.relationshipReferences.map(({ attribute, id, namespaceUri }) => ({ attribute, id, namespaceUri })),
        nativeRelationships: this.rootRelationships.map(({ id, type, target, targetMode }) => ({ id, type, target, targetMode })),
        nativeParts: this.parts.map((part) => ({ path: part.path, contentType: part.contentType, relationships: part.relationships.length })),
        embeddedWorkbook: this.oleWorkbook ? this._embeddedWorkbookRecord(true) : undefined,
        embeddedOfficePackage: this.oleOfficePackage ? this._embeddedOfficePackageRecord(true) : undefined,
        diagramText: this.diagramText,
        nativeChart: this._nativeChartBinding ? {
          titleLeaves: this._nativeChartTitleLeaves.length,
          title: this._nativeChartTitleLeaves.map((leaf) => leaf.text).join(""),
          dataPoints: this._nativeChartDataPoints.length,
        } : undefined,
        nativeTextLeaves: this._nativeTextRecords(),
        nativeLineLeaves: this._nativeLineRecords(),
        nativeStyleLeaves: this._nativeStyleRecords(),
        ...(this.text ? { text: this.text, textLength: this.textLength, ...(this.textTruncated ? { textTruncated: true } : {}) } : {}),
        deletionCapability: this.deletionCapability,
        placementCapability: this.placementCapability,
        bbox: [frame.left, frame.top, frame.width, frame.height],
        bboxUnit: "px",
        editable: false,
        editableFields,
      };
    }

    _embeddedWorkbookRecord(includeSourceSha256 = false) {
      const replacement = this._embeddedWorkbookReplacement;
      const part = replacement ? undefined : this.embeddedWorkbookPart();
      return {
        partPath: this.oleWorkbook.partPath,
        contentType: this.oleWorkbook.contentType,
        bytes: (replacement || part.bytes).length,
        ...(includeSourceSha256 ? { sourceSha256: this.oleWorkbook.sourceSha256 } : {}),
        replacementPending: Boolean(replacement),
      };
    }

    _embeddedOfficePackageRecord(includeSourceSha256 = false) {
      const replacement = this._embeddedOfficePackageReplacement;
      const part = replacement ? undefined : this.embeddedOfficePackagePart();
      return {
        kind: this.oleOfficePackage.kind,
        partPath: this.oleOfficePackage.partPath,
        contentType: this.oleOfficePackage.contentType,
        bytes: (replacement || part.bytes).length,
        ...(includeSourceSha256 ? { sourceSha256: this.oleOfficePackage.sourceSha256 } : {}),
        replacementPending: Boolean(replacement),
      };
    }

    layoutJson() {
      return {
        kind: "nativeObject",
        id: this.id,
        name: this.name,
        nativeKind: this.nativeKind,
        frame: this.position,
        relationships: this.rootRelationships.length,
        preservedParts: this.parts.length,
        embeddedWorkbook: this.oleWorkbook ? this._embeddedWorkbookRecord() : undefined,
        embeddedOfficePackage: this.oleOfficePackage ? this._embeddedOfficePackageRecord() : undefined,
        diagramText: this.diagramText,
        nativeChart: this._nativeChartCurrentRecord(),
        nativeTextLeaves: this._nativeTextRecords(),
        nativeLineLeaves: this._nativeLineRecords(),
        nativeStyleLeaves: this._nativeStyleRecords(),
        placementCapability: this.placementCapability,
        editable: false,
        editableFields: [
          ...(this.oleWorkbook ? ["embeddedWorkbook"] : []),
          ...(this.oleOfficePackage ? ["embeddedOfficePackage"] : []),
          ...(this._diagramTextBinding ? ["diagramText"] : []),
          ...(this._nativeChartTitleLeaves?.length ? ["chartTitleText"] : []),
          ...(this._nativeChartDataPoints?.length ? ["chartDataValue"] : []),
          ...(this._nativeTextLeaves?.length ? ["nativeText"] : []),
          ...nativeLineEditableFields(this._nativeLineLeaves),
          ...nativeStyleEditableFields(this._nativeStyleLeaves),
          ...(this.placementCapability.supported ? ["position"] : []),
        ],
      };
    }

    toSvg() {
      const p = this.position;
      if (!(p.width > 1 && p.height > 1)) return `<g data-native-object-id="${attrEscape(this.id)}" data-native-kind="${attrEscape(this.nativeKind)}"/>`;
      const label = this.name || this.nativeKind;
      return `<g data-native-object-id="${attrEscape(this.id)}" data-native-kind="${attrEscape(this.nativeKind)}"><rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#f8fafc" fill-opacity="0.72" stroke="#64748b" stroke-dasharray="6 4"/><text x="${p.left + 8}" y="${p.top + 20}" font-family="Arial" font-size="12" fill="#475569">${xmlEscape(label)}</text></g>`;
    }
  };
}
