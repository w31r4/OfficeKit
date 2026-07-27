import crypto from "node:crypto";
import { constants as FS_CONSTANTS } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "office-kit";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const ACCESSIBILITY_FIELDS = new Set(["title", "description"]);
const MAX_ACCESSIBILITY_TEXT_LENGTH = 1_024;
const MOVABLE_NAMESPACE_DECLARATIONS = new Map([
  ["xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main"],
  ["xmlns:p", "http://schemas.openxmlformats.org/presentationml/2006/main"],
  ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
]);
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("office-kit");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty string.");
  return value.trim();
}

function isXmlSafeText(value) {
  for (let index = 0; index < value.length; index += 1) {
    const codePoint = value.codePointAt(index);
    if (codePoint === undefined || codePoint === 0x7f ||
        !(
          codePoint === 0x9 || codePoint === 0xa || codePoint === 0xd ||
          (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
          (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
          (codePoint >= 0x10000 && codePoint <= 0x10ffff)
        )) return false;
    if (codePoint > 0xffff) index += 1;
  }
  return true;
}

function canonicalAccessibility(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new TypeError(label + " must be an object with optional title and description fields.");
  }
  const unsupported = Object.keys(value).filter((key) => !ACCESSIBILITY_FIELDS.has(key));
  if (unsupported.length) throw new TypeError(label + " has unsupported fields: " + unsupported.join(", ") + ".");
  const result = {};
  for (const field of ACCESSIBILITY_FIELDS) {
    if (!Object.hasOwn(value, field) || value[field] == null) continue;
    if (typeof value[field] !== "string" || !value[field] || value[field].length > MAX_ACCESSIBILITY_TEXT_LENGTH || !isXmlSafeText(value[field])) {
      throw new TypeError(`${label}.${field} must contain 1 through ${MAX_ACCESSIBILITY_TEXT_LENGTH} XML-safe characters.`);
    }
    result[field] = value[field];
  }
  return result;
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function withoutIdsAndAccessibility(value) {
  if (Array.isArray(value)) return value.map(withoutIdsAndAccessibility);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(Object.entries(value)
    .filter(([key]) => key !== "id" && key !== "accessibility")
    .map(([key, item]) => [key, withoutIdsAndAccessibility(item)]));
}

function semanticSnapshot(slide) {
  return withoutIdsAndAccessibility(slide.toProto());
}

async function modelSvgHashes(presentation) {
  return Promise.all(presentation.slides.items.map(async (slide) => {
    const render = await slide.export({ format: "svg" });
    const text = await render.text();
    if (!/<svg\b/i.test(text)) throw new Error("Presentation model render did not produce SVG.");
    return sha256(Buffer.from(text));
  }));
}

function xmlAttributes(tag) {
  const attributes = Object.create(null);
  for (const match of String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g)) attributes[match[1]] = match[3];
  return attributes;
}

function unescapeXml(value) {
  return String(value).replace(/&(#x[0-9a-f]+|#\d+|amp|apos|gt|lt|quot);/gi, (_, entity) => {
    const lower = entity.toLowerCase();
    if (lower === "amp") return "&";
    if (lower === "apos") return "'";
    if (lower === "gt") return ">";
    if (lower === "quot") return '"';
    const point = lower.startsWith("#x") ? Number.parseInt(lower.slice(2), 16) : Number.parseInt(lower.slice(1), 10);
    return Number.isSafeInteger(point) && point >= 0 && point <= 0x10ffff ? String.fromCodePoint(point) : "&" + entity + ";";
  });
}

function resolveRelationshipTarget(target) {
  const resolved = new URL(target, "https://officekit.invalid/ppt/presentation.xml");
  if (resolved.origin !== "https://officekit.invalid") throw new Error("Unexpected PPTX relationship target origin.");
  const partPath = resolved.pathname.replace(/^\/+/, "");
  if (!partPath.startsWith("ppt/") || partPath.split("/").includes("..")) throw new Error("Unsafe PPTX slide relationship target: " + JSON.stringify(target));
  return partPath;
}

async function orderedSlidePartPaths(zip) {
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const relationshipsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");
  if (!presentationXml || !relationshipsXml) throw new Error("PPTX is missing presentation.xml or its relationship part.");
  const relationships = new Map();
  for (const match of relationshipsXml.matchAll(/<Relationship\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type?.endsWith("/slide")) continue;
    if (attributes.TargetMode?.toLowerCase() === "external" || !attributes.Target) {
      throw new Error("Presentation slide relationship " + JSON.stringify(attributes.Id) + " is not an internal SlidePart.");
    }
    relationships.set(attributes.Id, resolveRelationshipTarget(attributes.Target));
  }
  const paths = [];
  for (const match of presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*>/gi)) {
    const target = relationships.get(xmlAttributes(match[0])["r:id"]);
    if (!target || !zip.file(target)) throw new Error("Presentation slide list has an unresolved SlidePart relationship.");
    paths.push(target);
  }
  if (!paths.length || new Set(paths).size !== paths.length) throw new Error("Presentation slide list must contain distinct, resolvable SlideParts.");
  return paths;
}

function targetCnvPr(xml, shapeName, partPath) {
  const matches = [...String(xml).matchAll(/<p:cNvPr\b[^>]*\/>/g)].filter((match) => unescapeXml(xmlAttributes(match[0]).name || "") === shapeName);
  if (matches.length !== 1) throw new Error(`Slide part ${partPath} must contain exactly one self-closing p:cNvPr named ${JSON.stringify(shapeName)}.`);
  const tag = matches[0][0];
  const attributes = xmlAttributes(tag);
  const unsupported = Object.keys(attributes).filter((key) => !new Set(["id", "name", "title", "descr", "hidden"]).has(key));
  if (unsupported.length || !/^\d+$/.test(attributes.id || "") || Number(attributes.id) < 1) {
    throw new Error("Selected p:cNvPr is not the bounded canonical title/description profile.");
  }
  const accessibility = {};
  if (attributes.title != null) accessibility.title = unescapeXml(attributes.title);
  if (attributes.descr != null) accessibility.description = unescapeXml(attributes.descr);
  return { tag, accessibility };
}

function canonicalizePresentationXmlForResidual(xml, label) {
  return String(xml).replace(/<[^>]+>/g, (tag) => {
    if (/^<\?/.test(tag) || /^<!/.test(tag) || /^<\//.test(tag)) return tag;
    const match = /^<([\w:.-]+)([\s\S]*?)(\/?)>$/.exec(tag);
    if (!match) throw new Error(`${label} contains unsupported XML markup during residual comparison.`);
    const [, name, sourceAttributes, slash] = match;
    let rest = sourceAttributes.trim();
    const attributes = [];
    while (rest) {
      const attribute = /^([:\w.-]+)=(["'])([\s\S]*?)\2\s*/.exec(rest);
      if (!attribute) throw new Error(`${label} contains unsupported XML attributes during residual comparison.`);
      const [, attributeName, , value] = attribute;
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

function stripTargetAccessibility(xml, shapeName, partPath) {
  const target = targetCnvPr(xml, shapeName, partPath);
  const stripped = target.tag
    .replace(/\s+title\s*=\s*(?:"[^"]*"|'[^']*')/, "")
    .replace(/\s+descr\s*=\s*(?:"[^"]*"|'[^']*')/, "");
  const withoutAccessibility = String(xml).replace(target.tag, stripped);
  // Open XML SDK can move the standard p/a/r declarations and reorder
  // attributes while saving one modified SlidePart. Canonicalize only that
  // XML-equivalent namespace placement and attribute order; every element,
  // non-namespace attribute value, and text node remains source-bound.
  return canonicalizePresentationXmlForResidual(withoutAccessibility, partPath);
}

async function assertPackageScope(sourceBytes, outputBytes, targetIndex, shapeName, expected, replacement) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const sourcePaths = Object.keys(sourceZip.files).sort();
  const outputPaths = Object.keys(outputZip.files).sort();
  if (!sameJson(sourcePaths, outputPaths)) throw new Error("Shape accessibility edit changed PPTX package topology.");
  const sourceSlidePaths = await orderedSlidePartPaths(sourceZip);
  const outputSlidePaths = await orderedSlidePartPaths(outputZip);
  if (!sameJson(sourceSlidePaths, outputSlidePaths)) throw new Error("Shape accessibility edit changed presentation slide-part routing.");
  if (!Number.isInteger(targetIndex) || targetIndex < 0 || targetIndex >= sourceSlidePaths.length) throw new Error("Resolved target slide index is outside the source PPTX slide list.");
  const targetPart = sourceSlidePaths[targetIndex];
  const sourceXml = await sourceZip.file(targetPart)?.async("text");
  const outputXml = await outputZip.file(targetPart)?.async("text");
  if (!sourceXml || !outputXml) throw new Error("Selected SlidePart is missing from source or output.");
  if (!sameJson(targetCnvPr(sourceXml, shapeName, targetPart).accessibility, expected)) {
    throw new Error("Selected source p:cNvPr does not match expectedAccessibility; no output was written.");
  }
  if (!sameJson(targetCnvPr(outputXml, shapeName, targetPart).accessibility, replacement)) {
    throw new Error("Saved p:cNvPr does not match replacementAccessibility.");
  }
  if (stripTargetAccessibility(sourceXml, shapeName, targetPart) !== stripTargetAccessibility(outputXml, shapeName, targetPart)) {
    throw new Error("Shape accessibility edit changed target SlidePart XML outside the selected p:cNvPr title/descr attributes.");
  }
  for (const partPath of sourcePaths) {
    if (sourceZip.files[partPath].dir || partPath === targetPart) continue;
    const before = await sourceZip.file(partPath).async("uint8array");
    const after = await outputZip.file(partPath).async("uint8array");
    if (!Buffer.from(before).equals(Buffer.from(after))) throw new Error("Shape accessibility edit changed non-target package part " + partPath + ".");
  }
  return { targetPart, partCount: sourcePaths.length, nonTargetPartsByteIdentical: true, targetResidualByteIdentical: true };
}

async function assertAbsent(filePath, label) {
  try {
    await fs.lstat(filePath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(label + " already exists; refusing to overwrite it.");
}

async function publishNoReplace(temporaryPath, finalPath, label) {
  try {
    await fs.link(temporaryPath, finalPath);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(label + " already exists; refusing to overwrite it.");
    if (!["EPERM", "EXDEV", "ENOTSUP", "EOPNOTSUPP"].includes(error?.code)) throw error;
    try {
      await fs.copyFile(temporaryPath, finalPath, FS_CONSTANTS.COPYFILE_EXCL);
    } catch (copyError) {
      if (copyError?.code === "EEXIST") throw new Error(label + " already exists; refusing to overwrite it.");
      throw copyError;
    }
  }
  await fs.rm(temporaryPath, { force: true }).catch(() => {});
}

export async function editPptxShapeAccessibility({ inputPath, outputPath, auditPath, slideName, shapeName, expectedAccessibility, replacementAccessibility }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const expectedSlideName = requiredText(slideName, "slideName");
  const expectedShapeName = requiredText(shapeName, "shapeName");
  const expected = canonicalAccessibility(expectedAccessibility, "expectedAccessibility");
  const replacement = canonicalAccessibility(replacementAccessibility, "replacementAccessibility");
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from source and PPTX output paths.");
  if (sameJson(expected, replacement)) throw new Error("replacementAccessibility must differ from expectedAccessibility.");
  await assertAbsent(finalPath, "outputPath");
  await assertAbsent(finalAuditPath, "auditPath");

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const slides = presentation.slides.items.filter((slide) => slide.name === expectedSlideName);
  if (slides.length !== 1) throw new Error("Expected exactly one imported slide named " + JSON.stringify(expectedSlideName) + "; found " + slides.length + ".");
  const targetSlide = slides[0];
  const targetIndex = presentation.slides.items.indexOf(targetSlide);
  const shapes = targetSlide.shapes.items.filter((shape) => shape.name === expectedShapeName);
  if (shapes.length !== 1) throw new Error("Expected exactly one top-level imported shape named " + JSON.stringify(expectedShapeName) + "; found " + shapes.length + ".");
  const targetShape = shapes[0];
  if (!sameJson(targetShape.accessibility || {}, expected)) {
    throw new Error("Selected imported shape does not match expectedAccessibility; no output was written.");
  }
  const sourceSemantics = presentation.slides.items.map(semanticSnapshot);
  const sourceRenderHashes = await modelSvgHashes(presentation);
  targetShape.setAccessibilityMetadata({
    title: Object.hasOwn(replacement, "title") ? replacement.title : null,
    description: Object.hasOwn(replacement, "description") ? replacement.description : null,
  });

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  let publishedOutput = false;
  let publishedAudit = false;
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertPackageScope(source, output, targetIndex, expectedShapeName, expected, replacement);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const reimportedSlides = reimported.slides.items.filter((slide) => slide.name === expectedSlideName);
    if (reimportedSlides.length !== 1) throw new Error("PPTX second import did not retain the unique target slide name.");
    const reimportedShapes = reimportedSlides[0].shapes.items.filter((shape) => shape.name === expectedShapeName);
    if (reimportedShapes.length !== 1 || !sameJson(reimportedShapes[0].accessibility || {}, replacement)) {
      throw new Error("PPTX second import did not retain replacement accessibility metadata.");
    }
    if (!sameJson(sourceSemantics, reimported.slides.items.map(semanticSnapshot))) {
      throw new Error("Shape accessibility edit changed non-accessibility presentation semantics.");
    }
    const outputRenderHashes = await modelSvgHashes(reimported);
    if (!sameJson(sourceRenderHashes, outputRenderHashes)) throw new Error("Shape accessibility edit changed a static slide model render.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    if (!Buffer.from(await fs.readFile(sourcePath)).equals(source)) throw new Error("Source PPTX changed during transaction; refusing to publish output.");
    const audit = {
      schema: "office-kit.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "source-bound-shape-accessibility-edit",
        slideName: expectedSlideName,
        slideNumber: targetIndex + 1,
        shapeName: expectedShapeName,
        partPath: packageScope.targetPart,
        nativeAttributes: ["p:cNvPr/@title", "p:cNvPr/@descr"],
        expectedAccessibility: expected,
        replacementAccessibility: replacement,
      },
      warnings: ["Static render verification proves visible-slide stability; final alternative-text review remains a native PowerPoint or LibreOffice accessibility review."],
      validation: {
        package: { ok: true, ...packageScope, onlyTargetSlidePartChanged: true },
        reimport: { ok: true, replacementSemanticsRetained: true },
        nonAccessibilitySemantics: { ok: true, stable: true },
        modelRender: { ok: true, sourceSha256: sourceRenderHashes, outputSha256: outputRenderHashes, byteIdentical: true },
        verify: { ok: verification.ok },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await publishNoReplace(temporaryPath, finalPath, "outputPath");
    publishedOutput = true;
    await publishNoReplace(temporaryAuditPath, finalAuditPath, "auditPath");
    publishedAudit = true;
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([
      fs.rm(temporaryPath, { force: true }),
      fs.rm(temporaryAuditPath, { force: true }),
      ...(publishedOutput ? [fs.rm(finalPath, { force: true })] : []),
      ...(publishedAudit ? [fs.rm(finalAuditPath, { force: true })] : []),
    ]);
    throw error;
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, slideName = "Accessibility metadata", shapeName = "decision-status", expected = '{"title":"Controlled rollout decision","description":"Status box explaining that the rollout is controlled."}', replacement = '{"title":"Go decision: controlled rollout"}'] = argv;
  if (!inputPath || !outputPath || !auditPath) {
    throw new Error("Usage: node officekit-shape-accessibility-edit-workflow.mjs <input.pptx> <output.pptx> <audit.json> [slideName] [shapeName] [expectedAccessibilityJson] [replacementAccessibilityJson]");
  }
  try {
    return { inputPath, outputPath, auditPath, slideName, shapeName, expectedAccessibility: JSON.parse(expected), replacementAccessibility: JSON.parse(replacement) };
  } catch (error) {
    throw new Error("expectedAccessibilityJson and replacementAccessibilityJson must be JSON objects: " + error.message);
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  editPptxShapeAccessibility(parseCli(process.argv.slice(2))).then((result) => {
    process.stdout.write(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, targetPart: result.audit.operation.partPath }) + "\n");
  }).catch((error) => {
    process.stderr.write((error?.stack || error?.message || String(error)) + "\n");
    process.exitCode = 1;
  });
}
