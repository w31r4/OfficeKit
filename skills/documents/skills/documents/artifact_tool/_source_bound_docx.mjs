import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";

import JSZip from "jszip";

export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const MOVABLE_NAMESPACE_DECLARATIONS = new Map([
  ["xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
  ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
  ["xmlns:wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"],
  ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
  ["xmlns:pic", "http://schemas.openxmlformats.org/drawingml/2006/picture"],
]);
const require = createRequire(import.meta.url);

export function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

export function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value;
}

export async function packageVersion() {
  const entry = require.resolve("office-kit");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

export async function assertAbsent(filePath, label) {
  try {
    await fs.lstat(filePath);
    throw new Error(`${label} already exists; refusing to overwrite it.`);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

export async function publishNoReplace(temporaryPath, finalPath) {
  await fs.link(temporaryPath, finalPath);
  await fs.rm(temporaryPath, { force: true });
}

export async function packageParts(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const parts = new Map();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    parts.set(name, Buffer.from(await entry.async("uint8array")));
  }
  return parts;
}

export async function changedParts(source, output, label = "Source-bound DOCX edit") {
  const [before, after] = await Promise.all([packageParts(source), packageParts(output)]);
  if (before.size !== after.size || [...before.keys()].some((name) => !after.has(name))) {
    throw new Error(`${label} changed the DOCX package part inventory.`);
  }
  return [...before.keys()].filter((name) => !before.get(name).equals(after.get(name))).sort();
}

export async function readPackagePartText(bytes, partPath, label = "DOCX package") {
  const zip = await JSZip.loadAsync(bytes);
  const entry = zip.file(partPath);
  if (!entry) throw new Error(`${label} has no ${partPath} part.`);
  return entry.async("text");
}

// Split only the immediate children of w:body. This deliberately does not
// pretend that nested table/textbox paragraphs are ordinary body blocks.
export function directBodyElements(xml, label = "DOCX document.xml") {
  const source = String(xml);
  const bodyMatch = /<w:body\b[^>]*>([\s\S]*)<\/w:body>/.exec(source);
  if (!bodyMatch) throw new Error(`${label} has no canonical w:body container.`);
  const bodyStart = (bodyMatch.index ?? 0) + bodyMatch[0].indexOf(">") + 1;
  const inner = bodyMatch[1];
  const elements = [];
  const stack = [];
  for (const match of inner.matchAll(/<\/?[\w:.-]+\b[^>]*>/g)) {
    const token = match[0];
    if (/^<\?/.test(token) || /^<!/.test(token)) throw new Error(`${label} has unsupported markup inside w:body.`);
    const closing = /^<\/([\w:.-]+)\s*>$/.exec(token);
    if (closing) {
      const current = stack.pop();
      if (!current || current.name !== closing[1]) throw new Error(`${label} has an unbalanced ${token} element.`);
      if (!stack.length) {
        elements.push({
          name: current.name,
          offset: bodyStart + current.offset,
          xml: inner.slice(current.offset, (match.index ?? 0) + token.length),
        });
      }
      continue;
    }
    const opening = /^<([\w:.-]+)\b[^>]*>$/.exec(token);
    if (!opening) throw new Error(`${label} has unsupported XML token ${token}.`);
    const selfClosing = /\/>$/.test(token);
    if (selfClosing) {
      if (!stack.length) elements.push({ name: opening[1], offset: bodyStart + (match.index ?? 0), xml: token });
      continue;
    }
    stack.push({ name: opening[1], offset: match.index ?? 0 });
  }
  if (stack.length) throw new Error(`${label} has an unclosed ${stack.at(-1).name} element.`);
  return elements;
}

// Open XML SDK may move a relationship namespace declaration from an individual
// element to the document root while changing an otherwise unrelated semantic
// leaf. Compare a strict tag/attribute canonical form so that namespace scope
// and attribute ordering do not masquerade as a document edit, while every
// non-namespace name, attribute value, text node, and element order stays bound.
export function canonicalizeXmlForResidual(xml, label) {
  return String(xml).replace(/<[^>]+>/g, (tag) => {
    if (/^<\?/.test(tag) || /^<!/.test(tag) || /^<\//.test(tag)) return tag;
    const match = /^<([\w:.-]+)([\s\S]*?)(\/?)>$/.exec(tag);
    if (!match) throw new Error(`${label} contains unsupported XML markup during residual comparison.`);
    const [, name, sourceAttributes, slash] = match;
    let rest = sourceAttributes.trim();
    const attributes = [];
    while (rest) {
      const attribute = /^([:\w.-]+)="([^"]*)"\s*/.exec(rest);
      if (!attribute) throw new Error(`${label} contains unsupported XML attributes during residual comparison.`);
      const [, attributeName, value] = attribute;
      if (MOVABLE_NAMESPACE_DECLARATIONS.has(attributeName)) {
        if (MOVABLE_NAMESPACE_DECLARATIONS.get(attributeName) !== value) {
          throw new Error(`${label} changes the ${attributeName} namespace binding.`);
        }
      } else {
        attributes.push([attributeName, value]);
      }
      rest = rest.slice(attribute[0].length);
    }
    attributes.sort(([left], [right]) => left.localeCompare(right));
    const suffix = attributes.length ? ` ${attributes.map(([attributeName, value]) => `${attributeName}="${value}"`).join(" ")}` : "";
    return `<${name}${suffix}${slash}>`;
  });
}
