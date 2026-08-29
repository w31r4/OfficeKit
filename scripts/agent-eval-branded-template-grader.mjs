import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import { PresentationFile } from "../src/presentation/index.mjs";
import { FileBlob } from "../src/shared/file-blob.mjs";
import { renderOfficeFile } from "./agent-eval-office-native-render.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";
import { BRANDED_TEMPLATE_FIXTURE } from "./agent-eval-branded-template.mjs";

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function check(id, category, passed, details = {}) {
  return { id, category, gate: false, passed: Boolean(passed), ...details };
}

function gate(id, category, passed, details = {}) {
  return { id, category, gate: true, passed: Boolean(passed), ...details };
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function xmlText(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
    .map((match) => match[1].replaceAll("&amp;", "&").replaceAll("&lt;", "<").replaceAll("&gt;", ">"))
    .join("");
}

function partPathFromTarget(sourcePart, target) {
  if (target.startsWith("/")) return target.slice(1);
  return path.posix.normalize(path.posix.join(path.posix.dirname(sourcePart), target));
}

async function zipPart(zip, partPath) {
  const entry = zip.file(partPath);
  return entry ? new Uint8Array(await entry.async("uint8array")) : null;
}

async function zipText(zip, partPath) {
  const entry = zip.file(partPath);
  return entry ? entry.async("text") : null;
}

function tableValues(xml) {
  const table = /<(?:[\w.-]+:)?tbl\b[\s\S]*?<\/(?:[\w.-]+:)?tbl>/.exec(String(xml))?.[0] || "";
  return [...table.matchAll(/<(?:[\w.-]+:)?tr\b[\s\S]*?<\/(?:[\w.-]+:)?tr>/g)].map((row) =>
    [...row[0].matchAll(/<(?:[\w.-]+:)?tc\b[\s\S]*?<\/(?:[\w.-]+:)?tc>/g)].map((cell) => xmlText(cell[0])));
}

function chartValues(xml) {
  const title = xmlText(/<(?:[\w.-]+:)?title\b[\s\S]*?<\/(?:[\w.-]+:)?title>/.exec(String(xml))?.[0] || "");
  const val = /<(?:[\w.-]+:)?val\b[\s\S]*?<\/(?:[\w.-]+:)?val>/.exec(String(xml))?.[0] || "";
  const values = [...val.matchAll(/<(?:[\w.-]+:)?pt\b[^>]*>[\s\S]*?<(?:[\w.-]+:)?v>([\s\S]*?)<\/(?:[\w.-]+:)?v>[\s\S]*?<\/(?:[\w.-]+:)?pt>/g)].map((match) => Number(match[1]));
  return { title, values };
}

async function slideImagePath(zip, slidePath) {
  const relsPath = `${path.posix.dirname(slidePath)}/_rels/${path.posix.basename(slidePath)}.rels`;
  const rels = await zipText(zip, relsPath) || "";
  const slideXml = await zipText(zip, slidePath) || "";
  const picture = /<p:pic\b[\s\S]*?name="product-image-target"[\s\S]*?<\/p:pic>/i.exec(slideXml)?.[0] || "";
  const id = /\br:embed="([^"]+)"/i.exec(picture)?.[1];
  if (!id) return null;
  for (const match of rels.matchAll(/<Relationship\b[^>]*>/gi)) {
    const attributes = Object.fromEntries([...match[0].matchAll(/\b([A-Za-z:]+)="([^"]*)"/g)].map((entry) => [entry[1], entry[2]]));
    if (attributes.Id === id && /\/image$/i.test(attributes.Type || "") && attributes.Target) return { id, path: partPathFromTarget(slidePath, attributes.Target) };
  }
  return null;
}

async function inspectBrandedTemplate(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const slidePaths = paths.filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name)).sort((left, right) => Number(left.match(/\d+/)?.[0]) - Number(right.match(/\d+/)?.[0]));
  const slideXml = await Promise.all(slidePaths.map((part) => zipText(zip, part)));
  const slide8Image = await slideImagePath(zip, "ppt/slides/slide8.xml");
  return {
    bytes,
    zip,
    paths,
    slidePaths,
    slideXml,
    title: xmlText([...String(slideXml[1] || "").matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)].find((match) => /name="board-title-target"/.test(match[0]))?.[0] || ""),
    table: tableValues(slideXml[4] || ""),
    chart: chartValues(await zipText(zip, "ppt/slides/charts/chart1.xml") || ""),
    slide8Image,
    smartArtPresent: BRANDED_TEMPLATE_FIXTURE.smartArtPartPaths.every((part) => paths.includes(part)) && /Brand operating SmartArt/.test(slideXml[3] || ""),
    olePresent: paths.includes(BRANDED_TEMPLATE_FIXTURE.oleWorkbookPath) && paths.includes(BRANDED_TEMPLATE_FIXTURE.olePreviewPath) && /Embedded finance workbook/.test(slideXml[5] || ""),
    notesPresent: paths.some((part) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(part)),
    commentsPresent: paths.includes("ppt/comments/modernComment.xml") && paths.includes("ppt/authors.xml"),
    transitionPresent: /<p:transition\b/.test(slideXml[2] || ""),
    customShowPresent: /Board route/.test(await zipText(zip, "ppt/presentation.xml") || ""),
  };
}

async function readAudit(workspace) {
  try { return JSON.parse(await fs.readFile(path.join(workspace, "outputs", "audit.json"), "utf8")); } catch { return null; }
}

function auditNoFallback(audit) {
  return audit?.provider?.silentFallback === false
    || (audit?.provider?.fallback_used === false && audit?.provider?.provider_switched !== true);
}

function auditRewrite(audit) {
  const strategy = audit?.savePolicy?.strategy;
  return typeof strategy === "string"
    && /rewrite/i.test(strategy)
    && audit.savePolicy.overwriteSource !== true
    && audit.savePolicy.sourceOverwriteAllowed !== true;
}

function auditSourceUnchanged(audit) {
  return audit?.validation?.sourceUnchanged === true
    || audit?.source?.immutable === true
    || audit?.source?.unchanged_after_delivery === true;
}

export async function gradeBrandedTemplateCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  if (item.id !== "pptx-branded-template-local-update") return { supported: false };
  const sourcePath = path.join(workspace, "inputs", "template.pptx");
  const outputPath = path.join(workspace, "outputs", "quarterly-board-updated.pptx");
  const replacementPath = path.join(workspace, "inputs", "replacement-product.png");
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const checks = [];
  let source;
  let output;
  let replacement;
  try {
    [source, output, replacement] = await Promise.all([
      inspectBrandedTemplate(sourcePath),
      inspectBrandedTemplate(outputPath),
      fs.readFile(replacementPath),
    ]);
  } catch (error) {
    checks.push(gate("branded-template-readable-output", "machine", false, { error: error.message }));
    checks.push(gate("branded-template-no-partial-success", "security", false, { error: error.message }));
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }

  checks.push(gate("branded-template-slide-count", "machine", source.slidePaths.length === 8 && output.slidePaths.length === 8, { source: source.slidePaths.length, output: output.slidePaths.length }));
  checks.push(gate("branded-template-title-edit", "machine", output.title === BRANDED_TEMPLATE_FIXTURE.replacementTitle && source.title === BRANDED_TEMPLATE_FIXTURE.originalTitle && !output.title.includes(BRANDED_TEMPLATE_FIXTURE.originalTitle), { source: source.title, output: output.title }));
  const sourceRevenue = source.table?.[1]?.slice(1, 4) || [];
  const outputRevenue = output.table?.[1]?.slice(1, 4) || [];
  checks.push(gate("branded-template-table-edit", "machine", JSON.stringify(sourceRevenue) === JSON.stringify(BRANDED_TEMPLATE_FIXTURE.tableOriginalValues) && JSON.stringify(outputRevenue) === JSON.stringify(BRANDED_TEMPLATE_FIXTURE.tableReplacementValues), { source: sourceRevenue, output: outputRevenue }));
  checks.push(gate("branded-template-chart-edit", "machine", source.chart.title === BRANDED_TEMPLATE_FIXTURE.chartOriginalTitle && JSON.stringify(source.chart.values) === JSON.stringify(BRANDED_TEMPLATE_FIXTURE.chartOriginalValues) && output.chart.title === BRANDED_TEMPLATE_FIXTURE.chartReplacementTitle && JSON.stringify(output.chart.values) === JSON.stringify(BRANDED_TEMPLATE_FIXTURE.chartReplacementValues), { source: source.chart, output: output.chart }));
  const outputImage = output.slide8Image?.path ? await zipPart(output.zip, output.slide8Image.path) : null;
  const sourceImage = source.slide8Image?.path ? await zipPart(source.zip, source.slide8Image.path) : null;
  checks.push(gate("branded-template-image-edit", "machine", Boolean(output.slide8Image?.path) && Buffer.from(outputImage || []).equals(Buffer.from(replacement)) && !Buffer.from(sourceImage || []).equals(Buffer.from(replacement)) && output.slideXml[7]?.includes(`descr="${BRANDED_TEMPLATE_FIXTURE.imageAlt}"`), { source: source.slide8Image, output: output.slide8Image }));
  checks.push(gate("branded-template-advanced-object-preservation", "security", ["smartArtPresent", "olePresent", "notesPresent", "commentsPresent", "transitionPresent", "customShowPresent"].every((key) => source[key] && output[key]), { source: Object.fromEntries(["smartArtPresent", "olePresent", "notesPresent", "commentsPresent", "transitionPresent", "customShowPresent"].map((key) => [key, source[key]])), output: Object.fromEntries(["smartArtPresent", "olePresent", "notesPresent", "commentsPresent", "transitionPresent", "customShowPresent"].map((key) => [key, output[key]])) }));
  const preservedParts = ["ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml", "ppt/slideMasters/theme/theme1.xml", ...BRANDED_TEMPLATE_FIXTURE.smartArtPartPaths, BRANDED_TEMPLATE_FIXTURE.oleWorkbookPath, BRANDED_TEMPLATE_FIXTURE.olePreviewPath, "ppt/notesSlides/notesSlide1.xml", "ppt/comments/modernComment.xml", "ppt/authors.xml", "ppt/presentation.xml", "ppt/slides/slide3.xml"];
  const preservedResults = [];
  for (const part of preservedParts) {
    const left = await zipPart(source.zip, part);
    const right = await zipPart(output.zip, part);
    preservedResults.push({ part, equal: Boolean(left && right && Buffer.from(left).equals(Buffer.from(right))) });
  }
  checks.push(gate("branded-template-preserved-parts", "security", preservedResults.every((entry) => entry.equal), { parts: preservedResults }));
  const sourceSet = new Set(source.paths);
  const outputSet = new Set(output.paths);
  const changedParts = [];
  for (const part of source.paths) {
    const left = await zipPart(source.zip, part);
    const right = await zipPart(output.zip, part);
    if (!right || !Buffer.from(left).equals(Buffer.from(right))) changedParts.push(part);
  }
  const outputOnlyParts = [...outputSet].filter((part) => !sourceSet.has(part));
  changedParts.push(...outputOnlyParts);
  const allowedParts = new Set(["ppt/slides/slide2.xml", "ppt/slides/slide5.xml", "ppt/slides/slide7.xml", "ppt/slides/charts/chart1.xml", "ppt/slides/slide8.xml", "ppt/slides/_rels/slide8.xml.rels", source.slide8Image?.path, output.slide8Image?.path].filter(Boolean));
  checks.push(gate("branded-template-only-target-parts", "security", outputOnlyParts.every((part) => part === output.slide8Image?.path) && changedParts.every((part) => allowedParts.has(part)), { changedParts, allowed: [...allowedParts] }));
  const sourceRender = await renderOfficeFile(sourcePath, "branded-template-source");
  const outputRender = await renderOfficeFile(outputPath, "branded-template-output");
  const pageStable = (index) => sourceRender.pages?.[index]?.pixelSha256 === outputRender.pages?.[index]?.pixelSha256;
  checks.push(gate("branded-template-eight-page-render", "visual", sourceRender.ok === true && outputRender.ok === true && sourceRender.pageCount === 8 && outputRender.pageCount === 8 && sourceRender.pages.every((page) => page.nonWhitePixels > 0) && outputRender.pages.every((page) => page.nonWhitePixels > 0), { source: sourceRender, output: outputRender }));
  checks.push(gate("branded-template-non-target-pages-stable", "visual", [0, 2, 3, 5].every(pageStable), { pages: [1, 3, 4, 6].map((page) => ({ page, stable: pageStable(page - 1) })) }));
  let reimport = null;
  try {
    reimport = await PresentationFile.importPptx(new FileBlob(output.bytes, { type: PPTX_MIME, name: "quarterly-board-updated.pptx" }));
  } catch {}
  checks.push(gate("branded-template-reimport", "machine", Boolean(reimport) && reimport.slides.count === 8, { slideCount: reimport?.slides?.count || null }));
  const publicWorkflowUsed = /officekit-branded-template-local-update-workflow\.mjs/i.test(commands);
  const typedRoundTrip = publicWorkflowUsed || (/PresentationFile\.importPptx/i.test(commands) && /PresentationFile\.exportPptx/i.test(commands));
  const noFallback = /silentFallback\s*[:=]\s*false|no[- ]fallback/i.test(`${commands}\n${finalMessage}`)
    || auditNoFallback(audit);
  checks.push(check("branded-template-trace:typed-roundtrip", "trace", typedRoundTrip, { expected: "PresentationFile.importPptx + PresentationFile.exportPptx or the public branded-template workflow" }));
  const forbiddenTrace = /direct\s+(?:xml|ooxml)|zip\/xml|flatten/i.test(commands);
  const structuredNoFlattening = audit?.validation?.noFlattening === true;
  checks.push(check("branded-template-trace:no-fallback", "trace", noFallback && (!forbiddenTrace || structuredNoFlattening), { expected: "typed OfficeKit workflow with no fallback" }));
  checks.push(gate("branded-template-audit", "trace", audit?.provider?.actual === "office-kit" && auditNoFallback(audit) && auditRewrite(audit) && audit?.source?.sha256 === sha256(source.bytes) && auditSourceUnchanged(audit), { audit: audit ? { provider: audit.provider, savePolicy: audit.savePolicy, source: audit.source, validation: audit.validation } : null }));
  const hardGatesPassed = checks.filter((entry) => entry.gate).every((entry) => entry.passed);
  const score = summarizeCaseScore(checks, item.grade, weights, hardGatesPassed);
  return { supported: true, graded: true, checks, evidence: { source: { sha256: sha256(source.bytes) }, output: { sha256: sha256(output.bytes) }, changedParts, sourceRender, outputRender }, pending: [], ...score };
}
