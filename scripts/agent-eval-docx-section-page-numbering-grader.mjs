import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import { DOCX_SECTION_PAGE_NUMBERING_FIXTURE } from "./agent-eval-office-fixtures.mjs";
import { renderOfficeFile } from "./agent-eval-office-native-render.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";

const PAGE_NUMBER_FORMATS = new Set(["decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"]);
const MOVABLE_NAMESPACE_DECLARATIONS = new Map([
  ["xmlns:w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"],
  ["xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
]);
const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const SHIPPED_SECTION_PAGE_NUMBERING_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/documents|node_modules\/office-kit\/skills\/documents\/skills\/documents)\/examples\/officekit-section-page-numbering-edit-workflow\.mjs(?:$|[\s"'`])/i;

function check(id, category, passed, details = {}) {
  return { id, category, gate: false, passed: Boolean(passed), ...details };
}

function gate(id, category, passed, details = {}) {
  return { id, category, gate: true, passed: Boolean(passed), ...details };
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function xmlAttributes(opening = "") {
  const result = {};
  for (const match of String(opening).matchAll(/([:\w.-]+)="([^"]*)"/g)) {
    result[match[1].split(":").at(-1)] = match[2];
  }
  return result;
}

function wordText(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
    .map((match) => match[1].replace(/<[^>]+>/g, "").replaceAll("&amp;", "&").replaceAll("&lt;", "<").replaceAll("&gt;", ">"))
    .join("");
}

function paragraphs(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?p>/g)].map((match) => wordText(match[1]));
}

function sectionProperties(xml) {
  return [...String(xml).matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((match, ordinal) => ({
    ordinal,
    xml: match[0],
    offset: match.index,
  }));
}

function canonicalizeXmlForResidual(xml, label) {
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

function parseCanonicalPageNumbering(sectionXml) {
  const leaves = [...String(sectionXml).matchAll(/<w:pgNumType\b[^>]*\/>/g)];
  if (!leaves.length) return { value: undefined, tag: null };
  if (leaves.length !== 1) throw new Error(`expected zero or one w:pgNumType leaf, found ${leaves.length}`);
  const tag = leaves[0][0];
  const attributes = xmlAttributes(tag);
  const unknown = Object.keys(attributes).filter((key) => key !== "start" && key !== "fmt");
  if (unknown.length) throw new Error(`unsupported w:pgNumType attributes: ${unknown.join(", ")}`);
  const value = {};
  if (Object.hasOwn(attributes, "start")) {
    if (!/^(?:0|[1-9]\d*)$/.test(attributes.start)) throw new Error("non-canonical w:start");
    const start = Number(attributes.start);
    if (!Number.isSafeInteger(start) || start > 2_147_483_647) throw new Error("out-of-range w:start");
    value.start = start;
  }
  if (Object.hasOwn(attributes, "fmt")) {
    if (!PAGE_NUMBER_FORMATS.has(attributes.fmt)) throw new Error(`unsupported w:fmt ${attributes.fmt}`);
    value.format = attributes.fmt;
  }
  if (!Object.keys(value).length) throw new Error("empty w:pgNumType");
  return { value, tag };
}

function inspectedSections(xml) {
  return sectionProperties(xml).map((section) => {
    try {
      const parsed = parseCanonicalPageNumbering(section.xml);
      return { ordinal: section.ordinal, xml: section.xml, pageNumbering: parsed.value, parseError: null };
    } catch (error) {
      return { ordinal: section.ordinal, xml: section.xml, pageNumbering: undefined, parseError: error.message };
    }
  });
}

function normalizeTargetPageNumberingXml(xml, sectionOrdinal) {
  const sections = sectionProperties(xml);
  const target = sections[sectionOrdinal];
  if (!target) return { ok: false, error: "target section is absent", normalized: null, pageNumbering: undefined, sectionCount: sections.length };
  try {
    const parsed = parseCanonicalPageNumbering(target.xml);
    if (!parsed.tag) return { ok: false, error: "target section has no w:pgNumType", normalized: null, pageNumbering: undefined, sectionCount: sections.length };
    const normalizedSection = target.xml.replace(parsed.tag, "<w:pgNumType officeKitMasked=\"true\"/>");
    return {
      ok: true,
      normalized: canonicalizeXmlForResidual(`${xml.slice(0, target.offset)}${normalizedSection}${xml.slice(target.offset + target.xml.length)}`, "target document XML"),
      pageNumbering: parsed.value,
      sectionCount: sections.length,
    };
  } catch (error) {
    return { ok: false, error: error.message, normalized: null, pageNumbering: undefined, sectionCount: sections.length };
  }
}

function canonicalSectionXml(section, label) {
  try {
    return section ? canonicalizeXmlForResidual(section.xml, label) : null;
  } catch {
    return null;
  }
}

export async function inspectSectionPageNumberingDocx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const entries = await Promise.all(paths.map(async (partPath) => [partPath, Buffer.from(await zip.file(partPath)?.async("uint8array") || [])]));
  const parts = new Map(entries);
  const documentPath = paths.find((name) => name.toLowerCase() === "word/document.xml") || null;
  const documentXml = documentPath ? (parts.get(documentPath) || Buffer.alloc(0)).toString("utf8") : "";
  const footerParts = paths.filter((name) => /^word\/footer\d+\.xml$/i.test(name)).map((partPath) => ({
    path: partPath,
    sha256: sha256(parts.get(partPath) || Buffer.alloc(0)),
    paragraphs: paragraphs((parts.get(partPath) || Buffer.alloc(0)).toString("utf8")),
  }));
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: Object.fromEntries(entries.map(([partPath, value]) => [partPath, sha256(value)])),
    documentPath,
    documentXml,
    bodyParagraphs: paragraphs(documentXml),
    sections: inspectedSections(documentXml),
    footerParts,
  };
}

function changedPackageParts(source, output) {
  if (!equalJson(source.paths, output.paths)) return null;
  return source.paths.filter((partPath) => source.partHashes[partPath] !== output.partHashes[partPath]);
}

function auditProvider(audit) {
  const provider = audit?.provider;
  return String(typeof provider === "string" ? provider : provider?.actual || provider?.selected || provider?.name || "");
}

function auditVersion(audit) {
  const provider = audit?.provider;
  return String(provider?.version || audit?.providerVersion || "");
}

function auditFallbackIsFalse(audit) {
  const provider = audit?.provider || {};
  const values = [provider.silentFallback, provider.silent_fallback, provider.fallbackUsed, provider.fallback_used, audit?.silentFallback, audit?.silent_fallback]
    .filter((value) => value !== undefined);
  return values.length > 0 && values.every((value) => value === false || value === "false");
}

function auditStrategy(audit) {
  const policy = audit?.savePolicy || audit?.save_policy || audit?.saveStrategy || audit?.save_strategy;
  return String(typeof policy === "string" ? policy : policy?.strategy || policy?.selected || audit?.strategy || "");
}

function auditHash(audit, side) {
  const record = audit?.[side] || {};
  return String(record.sha256 || audit?.[side + "Sha256"] || audit?.[side + "_sha256"] || "");
}

function auditOperation(audit) {
  const operation = audit?.operation;
  return String(typeof operation === "string" ? operation : operation?.type || operation?.name || "");
}

function visiblePageNumberingChange(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount;
  const targetChanged = pageCountsMatch
    && source?.pages?.[0]?.pixelSha256 !== output?.pages?.[0]?.pixelSha256;
  const nonTargetPagesStable = pageCountsMatch
    && source?.pages?.length === output?.pages?.length
    && source.pages.slice(1).every((page, index) => page.pixelSha256 === output.pages[index + 1]?.pixelSha256);
  return { available, rendered, pageCountsMatch, targetChanged, nonTargetPagesStable };
}

function usedTypedSectionPageNumberingRoundTrip(commandText) {
  const directPublicApi = /(?:DocumentFile\.)?importDocx/i.test(commandText)
    && /(?:DocumentFile\.)?exportDocx/i.test(commandText);
  return directPublicApi || SHIPPED_SECTION_PAGE_NUMBERING_WORKFLOW.test(commandText);
}

export function gradeDocxSectionPageNumberingEvidence({ evidence, audit, commands }) {
  const fixture = DOCX_SECTION_PAGE_NUMBERING_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const target = fixture.target;
  const sourceTarget = source.sections[target.sectionOrdinal];
  const outputTarget = output.sections[target.sectionOrdinal];
  const sourceSibling = source.sections[fixture.sibling.sectionOrdinal];
  const outputSibling = output.sections[fixture.sibling.sectionOrdinal];
  const sourceTerminal = source.sections.at(-1);
  const outputTerminal = output.sections.at(-1);
  const siblingXmlStable = canonicalSectionXml(sourceSibling, "source sibling section") === canonicalSectionXml(outputSibling, "output sibling section");
  const terminalXmlStable = canonicalSectionXml(sourceTerminal, "source terminal section") === canonicalSectionXml(outputTerminal, "output terminal section");
  const sourceResidual = normalizeTargetPageNumberingXml(source.documentXml, target.sectionOrdinal);
  const outputResidual = normalizeTargetPageNumberingXml(output.documentXml, target.sectionOrdinal);
  const residualStable = sourceResidual.ok && outputResidual.ok
    && sourceResidual.sectionCount === outputResidual.sectionCount
    && sourceResidual.normalized === outputResidual.normalized;
  const changed = changedPackageParts(source, output);
  const footerPartsStable = source.footerParts.length === output.footerParts.length
    && source.footerParts.every((part, index) => part.path === output.footerParts[index]?.path && part.sha256 === output.footerParts[index]?.sha256);
  const visual = visiblePageNumberingChange(evidence.visual?.source, evidence.visual?.output);
  const commandText = commands.join("\n");
  const auditTarget = audit?.operation?.target || {};
  return [
    check("docx-page-numbering-machine:fixture-canonical-profile", "machine", source.sections.length === fixture.nativeSectionCount
      && equalJson(sourceTarget?.pageNumbering, target.originalPageNumbering)
      && equalJson(sourceSibling?.pageNumbering, fixture.sibling.pageNumbering)
      && sourceTerminal?.pageNumbering === undefined
      && source.sections.every((section) => section.parseError === null)
      && source.footerParts.length === fixture.footerSectionCount
      && source.bodyParagraphs.includes(fixture.title)
      && fixture.body.every((paragraph) => source.bodyParagraphs.includes(paragraph)), {
      sections: source.sections.map((section) => ({ ordinal: section.ordinal, pageNumbering: section.pageNumbering, parseError: section.parseError })),
      footerParts: source.footerParts.map((part) => ({ path: part.path, paragraphs: part.paragraphs })),
      bodyParagraphs: source.bodyParagraphs,
    }),
    check("docx-page-numbering-machine:requested-section-value-edited", "machine", equalJson(outputTarget?.pageNumbering, target.replacementPageNumbering)
      && equalJson(outputSibling?.pageNumbering, fixture.sibling.pageNumbering)
      && outputTerminal?.pageNumbering === undefined
      && output.sections.every((section) => section.parseError === null), {
      sourceTarget: sourceTarget?.pageNumbering,
      outputTarget: outputTarget?.pageNumbering,
      sourceSibling: sourceSibling?.pageNumbering,
      outputSibling: outputSibling?.pageNumbering,
      sourceTerminal: sourceTerminal?.pageNumbering,
      outputTerminal: outputTerminal?.pageNumbering,
    }),
    check("docx-page-numbering-machine:document-xml-residual-stable", "machine", residualStable, {
      source: { ok: sourceResidual.ok, error: sourceResidual.error || null, pageNumbering: sourceResidual.pageNumbering },
      output: { ok: outputResidual.ok, error: outputResidual.error || null, pageNumbering: outputResidual.pageNumbering },
    }),
    check("docx-page-numbering-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("docx-page-numbering-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("docx-page-numbering-visual:target-page-number-format-visible", "visual", visual.targetChanged && visual.nonTargetPagesStable, {
      visual: evidence.visual,
      note: "The fixture places a PAGE field on each native section. Only the first rendered page may change when its lower-Roman format becomes decimal.",
    }),
    gate("docx-page-numbering-security:only-document-part-changed", "security", equalJson(changed, ["word/document.xml"]), {
      changed,
    }),
    gate("docx-page-numbering-security:sibling-sections-footers-and-inventory-preserved", "security", footerPartsStable
      && equalJson(changed, ["word/document.xml"])
      && siblingXmlStable
      && terminalXmlStable, {
      footerPartsStable,
      sourcePaths: source.paths,
      outputPaths: output.paths,
      siblingXmlStable,
      terminalXmlStable,
    }),
    gate("docx-page-numbering-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("docx-page-numbering-trace:office-kit-provider", "trace", /office[- ]?kit/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("docx-page-numbering-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("docx-page-numbering-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("docx-page-numbering-trace:bounded-operation", "trace", /section.*page.*number|page.*number.*section/i.test(auditOperation(audit))
      && auditTarget.blockIndex === target.blockIndex
      && auditTarget.sectionOrdinal === target.sectionOrdinal
      && equalJson(audit?.operation?.sourcePageNumbering, target.originalPageNumbering)
      && equalJson(audit?.operation?.replacementPageNumbering, target.replacementPageNumbering), {
      operation: audit?.operation || null,
    }),
    check("docx-page-numbering-trace:typed-roundtrip", "trace", usedTypedSectionPageNumberingRoundTrip(commandText), {
      expected: "public DocumentFile importDocx/exportDocx calls or the integrity-protected published section page-numbering workflow",
    }),
    check("docx-page-numbering-trace:second-import", "trace", audit?.validation?.reimport?.ok === true
      && audit?.validation?.reimport?.sectionBlockIndex === target.blockIndex
      && audit?.validation?.reimport?.sectionOrdinal === target.sectionOrdinal
      && equalJson(audit?.validation?.reimport?.pageNumbering, target.replacementPageNumbering), {
      validation: audit?.validation || null,
    }),
  ];
}

async function readAudit(workspace) {
  try {
    return JSON.parse(await fs.readFile(path.join(workspace, "outputs", "audit.json"), "utf8"));
  } catch {
    return null;
  }
}

export async function gradeDocxSectionPageNumberingCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const fixture = DOCX_SECTION_PAGE_NUMBERING_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const sourcePath = path.join(workspace, "inputs", fixture.documentName);
  const outputPath = path.join(workspace, "outputs", "front-matter-page-numbering-reviewed.docx");
  let source;
  let output;
  try {
    [source, output] = await Promise.all([
      inspectSectionPageNumberingDocx(sourcePath),
      inspectSectionPageNumberingDocx(outputPath),
    ]);
  } catch (error) {
    const checks = [
      gate("docx-page-numbering-machine:readable-output", "machine", false, { error: error.message }),
      gate("docx-page-numbering-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }

  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(sourcePath, "docx-page-numbering-source"),
    renderOfficeFile(outputPath, "docx-page-numbering-output"),
  ]);
  const visualUnavailable = [sourceRender, outputRender].find((result) => !result.available);
  if (visualUnavailable) {
    return {
      supported: true,
      graded: false,
      checks: [],
      evidence: { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage },
      pending: ["native LibreOffice/Poppler document rendering"],
      infrastructureErrors: [visualUnavailable.reason],
    };
  }
  const evidence = { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage };
  const checks = gradeDocxSectionPageNumberingEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}
