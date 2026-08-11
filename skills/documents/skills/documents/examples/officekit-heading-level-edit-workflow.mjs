import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { DocumentFile, FileBlob } from "office-kit";
import {
  DOCX_MIME,
  assertAbsent,
  canonicalizeXmlForResidual,
  changedParts,
  directBodyElements,
  packageVersion,
  publishNoReplace,
  readPackagePartText,
  requiredText,
  sha256,
} from "../artifact_tool/_source_bound_docx.mjs";

const MAX_TEXT_CHARS = 1_000_000;

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function jsonClone(value) {
  return JSON.parse(JSON.stringify(value));
}

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return index;
}

function headingLevel(value, label, { allowBody = false } = {}) {
  const level = Number(value);
  const minimum = allowBody ? 0 : 1;
  if (!Number.isSafeInteger(level) || level < minimum || level > 9) {
    throw new TypeError(`${label} must be an integer from ${minimum} through 9.`);
  }
  return level;
}

function directOutlineLevel(value, label) {
  if (value == null || value === "inherit") return null;
  const level = Number(value);
  if (!Number.isSafeInteger(level) || level < 0 || level > 9) {
    throw new TypeError(`${label} must be null/inherit or an integer from 0 through 9.`);
  }
  return level;
}

function exactText(value, label, { visible = false } = {}) {
  if (typeof value !== "string" || value.length > MAX_TEXT_CHARS || /\u0000/u.test(value)) {
    throw new TypeError(`${label} must be a string of at most ${MAX_TEXT_CHARS} characters without NUL.`);
  }
  if (visible && !value.trim()) throw new TypeError(`${label} must contain visible non-whitespace text.`);
  return value;
}

function canonicalOutlineLeaf(markup, label) {
  const match = /^<w:outlineLvl\s+w:val="(0|[1-9])"\s*\/>$/.exec(markup);
  if (!match) throw new Error(`${label} must be a canonical self-closing w:outlineLvl with exactly one w:val attribute.`);
  return Number(match[1]);
}

function directOutlineLeaves(inner, label) {
  const leaves = [];
  const stack = [];
  for (const match of String(inner).matchAll(/<\/?[\w:.-]+\b[^>]*>/g)) {
    const markup = match[0];
    const closing = /^<\/([\w:.-]+)\s*>$/.exec(markup);
    if (closing) {
      const current = stack.pop();
      if (!current || current !== closing[1]) throw new Error(`${label} has unbalanced paragraph-property markup.`);
      continue;
    }
    const opening = /^<([\w:.-]+)\b[^>]*>$/.exec(markup);
    if (!opening) throw new Error(`${label} contains unsupported paragraph-property markup.`);
    const name = opening[1];
    const selfClosing = /\/\s*>$/.test(markup);
    if (name === "w:outlineLvl") {
      if (stack.length || !selfClosing) throw new Error(`${label} has nested or non-leaf w:outlineLvl markup.`);
      leaves.push({ markup, offset: match.index ?? 0, level: canonicalOutlineLeaf(markup, `${label} w:outlineLvl`) });
    }
    if (!selfClosing) stack.push(name);
  }
  if (stack.length) throw new Error(`${label} has unclosed paragraph-property markup.`);
  return leaves;
}

function paragraphOutlineProfile(paragraphXml, label) {
  const source = String(paragraphXml);
  const properties = [...source.matchAll(/<w:pPr\b[^>]*\/>|<w:pPr\b[^>]*>[\s\S]*?<\/w:pPr>/g)];
  if (properties.length > 1) throw new Error(`${label} has multiple w:pPr containers.`);
  if (!properties.length) return { level: null, masked: source };
  const propertyMarkup = properties[0][0];
  const propertyOffset = properties[0].index ?? 0;
  const selfClosing = /^<w:pPr\b([^>]*)\/>$/.exec(propertyMarkup);
  const container = /^<w:pPr\b([^>]*)>([\s\S]*)<\/w:pPr>$/.exec(propertyMarkup);
  const attributes = selfClosing?.[1] ?? container?.[1];
  if (attributes === undefined) throw new Error(`${label} has unsupported w:pPr markup.`);
  const inner = selfClosing ? "" : container?.[2] ?? "";
  const leaves = directOutlineLeaves(inner, label);
  if (leaves.length > 1) throw new Error(`${label} has duplicate w:outlineLvl leaves.`);
  const leaf = leaves[0];
  const maskedInner = leaf ? `${inner.slice(0, leaf.offset)}${inner.slice(leaf.offset + leaf.markup.length)}` : inner;
  const maskedProperties = !attributes.trim() && !maskedInner.trim()
    ? ""
    : `<w:pPr${attributes}>${maskedInner}</w:pPr>`;
  return {
    level: leaf?.level ?? null,
    masked: `${source.slice(0, propertyOffset)}${maskedProperties}${source.slice(propertyOffset + propertyMarkup.length)}`,
  };
}

function normalizedTargetParagraphXml(xml, blockIndex, expectedBlockCount, expectedDirectLevel, label) {
  const blocks = directBodyElements(xml, label).filter((element) => element.name !== "w:sectPr");
  if (blocks.length !== expectedBlockCount) {
    throw new Error(`${label} has ${blocks.length} direct native blocks, but import exposed ${expectedBlockCount}; refusing an ambiguous source-to-model mapping.`);
  }
  const target = blocks[blockIndex];
  if (!target || target.name !== "w:p") throw new Error(`${label} block ${blockIndex} is not one direct w:p paragraph.`);
  const profile = paragraphOutlineProfile(target.xml, label);
  if (profile.level !== expectedDirectLevel) {
    throw new Error(`${label} direct outline level does not match the bound source value: expected ${JSON.stringify(expectedDirectLevel)}, observed ${JSON.stringify(profile.level)}.`);
  }
  return {
    directOutlineLevel: profile.level,
    normalized: canonicalizeXmlForResidual(
      `${String(xml).slice(0, target.offset)}${profile.masked}${String(xml).slice(target.offset + target.xml.length)}`,
      label,
    ),
  };
}

function directModelOutlineLevel(block) {
  return Object.hasOwn(block.paragraphFormat || {}, "outlineLevel") ? block.paragraphFormat.outlineLevel : null;
}

function paragraphSnapshot(block, blockIndex) {
  return {
    id: block.id,
    blockIndex,
    name: block.name || "",
    styleId: block.styleId || "",
    text: String(block.text || ""),
    textEditable: block.textEditable === true,
    textPatchable: block.textPatchable === true,
    pendingTextPatches: block.textPatches?.length || 0,
    directOutlineLevel: directModelOutlineLevel(block),
    paragraphFormat: jsonClone(block.paragraphFormat || {}),
    runs: jsonClone(block.runs || []),
  };
}

function selectParagraph(document, {
  headingBlockIndex,
  expectedText,
  expectedStyleId,
  expectedDirectOutlineLevel,
}) {
  const blockIndex = boundedIndex(headingBlockIndex, "headingBlockIndex");
  const block = document.blocks[blockIndex];
  if (!block || block.kind !== "paragraph") throw new Error("headingBlockIndex does not identify an imported paragraph block.");
  if (document.resolve(block.id) !== block) throw new Error("Selected heading locator did not resolve to the inspected paragraph.");
  if (block.textEditable !== true || block.textPatches?.length) {
    throw new Error("Selected heading is not one safely editable modeled source paragraph.");
  }
  if (block.text !== expectedText) {
    throw new Error(`Selected heading text does not match the expected source value: expected ${JSON.stringify(expectedText)}, observed ${JSON.stringify(block.text)}.`);
  }
  if ((block.styleId || "") !== expectedStyleId) {
    throw new Error(`Selected heading style does not match the expected source value: expected ${JSON.stringify(expectedStyleId)}, observed ${JSON.stringify(block.styleId || "")}.`);
  }
  const actualDirectLevel = directModelOutlineLevel(block);
  if (actualDirectLevel !== expectedDirectOutlineLevel) {
    throw new Error(`Selected heading direct outline level does not match the expected source value: expected ${JSON.stringify(expectedDirectOutlineLevel)}, observed ${JSON.stringify(actualDirectLevel)}.`);
  }
  return { block, blockIndex, snapshot: paragraphSnapshot(block, blockIndex) };
}

function selectedHeadingIssue(report, selected, expectedHeadingLevel, expectedPreviousHeadingLevel) {
  const matches = report.issues.filter((issue) =>
    issue.type === "headingLevelSkipped" &&
    issue.id === selected.snapshot.id &&
    issue.blockIndex === selected.blockIndex);
  if (matches.length !== 1) throw new Error("Selected paragraph is not exactly one headingLevelSkipped accessibility issue.");
  const issue = matches[0];
  if (issue.headingLevel !== expectedHeadingLevel || issue.previousHeadingLevel !== expectedPreviousHeadingLevel) {
    throw new Error(`Selected heading issue does not match the expected source levels: expected ${expectedPreviousHeadingLevel} -> ${expectedHeadingLevel}, observed ${issue.previousHeadingLevel} -> ${issue.headingLevel}.`);
  }
  return jsonClone(issue);
}

function issueKey(issue) {
  return `${issue.type}\u0000${issue.id}\u0000${issue.blockIndex}`;
}

function assertAccessibilityImproved(before, after, selectedIssue) {
  if (after.issues.some((issue) => issueKey(issue) === issueKey(selectedIssue))) {
    throw new Error("Replacement outline level did not resolve the selected headingLevelSkipped issue.");
  }
  const beforeKeys = new Set(before.issues.map(issueKey));
  const newIssues = after.issues.filter((issue) => !beforeKeys.has(issueKey(issue)));
  if (newIssues.length) {
    throw new Error(`Replacement outline level introduces new accessibility machine issues: ${newIssues.map(issueKey).join(", ")}.`);
  }
  if (after.summary.headingLevelSkips >= before.summary.headingLevelSkips) {
    throw new Error("Replacement outline level did not reduce the document heading-level skip count.");
  }
}

function maskedDocumentProjection(document, targetId) {
  const projection = jsonClone(document.toProto());
  const target = projection.blocks.find((block) => block.id === targetId);
  if (!target || target.kind !== "paragraph") throw new Error("Selected heading is missing from the document semantic projection.");
  if (target.paragraphFormat) {
    delete target.paragraphFormat.outlineLevel;
    if (!Object.keys(target.paragraphFormat).length) delete target.paragraphFormat;
  }
  return projection;
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/iu.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length, sha256: sha256(Buffer.from(svg, "utf8")) };
}

/**
 * Repairs one modeled imported heading-level skip by setting one canonical
 * direct w:outlineLvl value. The paragraph style, visible text/runs, package
 * graph, and every other modeled semantic remain source-owned.
 */
export async function editImportedHeadingLevel({
  inputPath,
  outputPath,
  auditPath,
  headingBlockIndex,
  expectedText,
  expectedStyleId,
  expectedDirectOutlineLevel,
  expectedHeadingLevel,
  expectedPreviousHeadingLevel,
  replacementHeadingLevel,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const blockIndex = boundedIndex(headingBlockIndex, "headingBlockIndex");
  const text = exactText(expectedText, "expectedText", { visible: true });
  const styleId = exactText(expectedStyleId, "expectedStyleId");
  const expectedDirect = directOutlineLevel(expectedDirectOutlineLevel, "expectedDirectOutlineLevel");
  const expectedLevel = headingLevel(expectedHeadingLevel, "expectedHeadingLevel");
  const expectedPrevious = headingLevel(expectedPreviousHeadingLevel, "expectedPreviousHeadingLevel", { allowBody: true });
  const replacementLevel = headingLevel(replacementHeadingLevel, "replacementHeadingLevel");
  if (replacementLevel > expectedPrevious + 1) {
    throw new Error(`replacementHeadingLevel must be at most ${expectedPrevious + 1} after the bound preceding heading.`);
  }
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectParagraph(document, {
    headingBlockIndex: blockIndex,
    expectedText: text,
    expectedStyleId: styleId,
    expectedDirectOutlineLevel: expectedDirect,
  });
  const sourceAccessibility = document.auditAccessibility({ maxChars: 200_000 });
  const sourceIssue = selectedHeadingIssue(sourceAccessibility, selected, expectedLevel, expectedPrevious);
  const sourceProjection = maskedDocumentProjection(document, selected.snapshot.id);
  const sourceRender = await modelRender(document);
  const sourceXml = await readPackagePartText(source, "word/document.xml", "Source DOCX package");
  const sourceResidual = normalizedTargetParagraphXml(
    sourceXml,
    selected.blockIndex,
    document.blocks.length,
    expectedDirect,
    "source target heading",
  );

  selected.block.paragraphFormat = { ...(selected.block.paragraphFormat || {}), outlineLevel: replacementLevel - 1 };
  const updatedAccessibility = document.auditAccessibility({ maxChars: 200_000 });
  assertAccessibilityImproved(sourceAccessibility, updatedAccessibility, sourceIssue);
  const sourceVerification = document.verify({ visualQa: true });
  if (!sourceVerification.ok) throw new Error(`Replacement heading level fails document verification: ${sourceVerification.ndjson}`);
  const updatedRender = await modelRender(document);
  if (updatedRender.sha256 !== sourceRender.sha256) {
    throw new Error("Heading outline-level edit changed the model visual projection.");
  }

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output, "Source-bound heading-level edit");
    if (!equalJson(changed, ["word/document.xml"])) {
      throw new Error(`Source-bound heading-level edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }
    const outputXml = await readPackagePartText(output, "word/document.xml", "Output DOCX package");
    const outputResidual = normalizedTargetParagraphXml(
      outputXml,
      selected.blockIndex,
      document.blocks.length,
      replacementLevel - 1,
      "output target heading",
    );
    if (outputResidual.normalized !== sourceResidual.normalized) {
      throw new Error("Heading-level edit changed word/document.xml outside the bound w:outlineLvl leaf and its necessary empty w:pPr container.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectParagraph(reimported, {
      headingBlockIndex: selected.blockIndex,
      expectedText: text,
      expectedStyleId: styleId,
      expectedDirectOutlineLevel: replacementLevel - 1,
    });
    if (roundTrip.snapshot.id !== selected.snapshot.id) throw new Error("Second import did not preserve the selected heading identity.");
    if (!equalJson(maskedDocumentProjection(reimported, roundTrip.snapshot.id), sourceProjection)) {
      throw new Error("DOCX export changed modeled document semantics outside the requested direct outline level.");
    }
    const accessibility = reimported.auditAccessibility({ maxChars: 200_000 });
    assertAccessibilityImproved(sourceAccessibility, accessibility, sourceIssue);
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
    const render = await modelRender(reimported);
    if (render.sha256 !== sourceRender.sha256) throw new Error("Second import changed the model visual projection.");

    const audit = {
      schema: "office-kit.docx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "office-kit", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", noReplace: true },
      operation: {
        type: "source-bound-heading-level-edit",
        target: { id: selected.snapshot.id, blockIndex: selected.blockIndex },
        expected: {
          text,
          styleId,
          directOutlineLevel: expectedDirect,
          headingLevel: expectedLevel,
          previousHeadingLevel: expectedPrevious,
        },
        replacement: { directOutlineLevel: replacementLevel - 1, headingLevel: replacementLevel },
      },
      validation: {
        changedParts: changed,
        headingLevelXmlResidual: {
          ok: true,
          blockIndex: selected.blockIndex,
          normalizedSha256: sha256(Buffer.from(sourceResidual.normalized, "utf8")),
        },
        reimport: {
          ok: true,
          headingId: roundTrip.snapshot.id,
          blockIndex: roundTrip.blockIndex,
          styleId,
          directOutlineLevel: replacementLevel - 1,
          headingLevel: replacementLevel,
        },
        accessibility: {
          selectedIssueResolved: true,
          headingLevelSkipsBefore: sourceAccessibility.summary.headingLevelSkips,
          headingLevelSkipsAfter: accessibility.summary.headingLevelSkips,
          machineCheckPassed: accessibility.machineCheckPassed,
          conformanceClaimed: false,
          manualReviewRequired: accessibility.manualReviewRequired,
        },
        verify: { ok: true },
        modelRender: { ok: true, unchanged: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["This transaction changes one direct Word outline level without changing the paragraph style or visible formatting. Confirm the intended document hierarchy and inspect a native render; this is not Word Accessibility Checker or WCAG conformance."],
    };
    await fs.writeFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`, { flag: "wx" });
    await publishNoReplace(temporaryPath, finalPath);
    try {
      await publishNoReplace(temporaryAuditPath, finalAuditPath);
    } catch (error) {
      await fs.rm(finalPath, { force: true });
      throw error;
    }
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

export function parseHeadingLevelEditCli(argv) {
  const [
    inputPath,
    outputPath,
    auditPath,
    headingBlockIndex,
    expectedText,
    expectedStyleId,
    expectedDirectOutlineLevel,
    expectedHeadingLevel,
    expectedPreviousHeadingLevel,
    replacementHeadingLevel,
  ] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    headingBlockIndex: boundedIndex(headingBlockIndex, "headingBlockIndex"),
    expectedText: exactText(expectedText, "expectedText", { visible: true }),
    expectedStyleId: exactText(expectedStyleId, "expectedStyleId"),
    expectedDirectOutlineLevel: directOutlineLevel(expectedDirectOutlineLevel, "expectedDirectOutlineLevel"),
    expectedHeadingLevel: headingLevel(expectedHeadingLevel, "expectedHeadingLevel"),
    expectedPreviousHeadingLevel: headingLevel(expectedPreviousHeadingLevel, "expectedPreviousHeadingLevel", { allowBody: true }),
    replacementHeadingLevel: headingLevel(replacementHeadingLevel, "replacementHeadingLevel"),
  };
}

export function headingLevelCliOutput(result) {
  return {
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
    headingLevel: result.audit.operation.replacement.headingLevel,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedHeadingLevel(parseHeadingLevelEditCli(process.argv.slice(2)));
  console.log(JSON.stringify(headingLevelCliOutput(result)));
}
