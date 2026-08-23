#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { constants } from "node:fs";
import {
  chmod,
  copyFile,
  mkdir,
  readFile,
  stat,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const defaultDefinitions = path.join(repoRoot, "evals/pptx-programmable-import/source-derived-companion.v1.json");
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const REPLACEMENT_PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.worker) return runWorker(args);

  const definitionsPath = path.resolve(args.definitions || defaultDefinitions);
  const definitionsBytes = await readFile(definitionsPath);
  const definitions = JSON.parse(definitionsBytes);
  validateDefinitions(definitions);
  const assetsDir = path.resolve(args["assets-dir"] || process.env[definitions.assetsEnvironment] || "/Users/zfang/Downloads/飞书20260814-175228");
  const runRoot = path.resolve(required(args, "run-root"));
  await requireAbsent(runRoot, "run root");
  await mkdir(runRoot, { recursive: true });
  const packageRoot = path.resolve(args["package-root"] || repoRoot);
  const packageMetadata = JSON.parse(await readFile(path.join(packageRoot, "package.json"), "utf8"));
  if (packageMetadata.name !== "office-kit") throw new Error(`--package-root is not office-kit: ${packageRoot}`);
  const officekitBin = path.join(packageRoot, packageMetadata.bin?.officekit || "bin/officekit.mjs");
  await stat(officekitBin);
  const repetitions = args.repetitions
    ? positiveInteger(args.repetitions, "repetitions")
    : definitions.repetitionsPerCase;
  const render = args["no-render"] !== true;
  const targetRenderer = parseTargetRenderer(args["target-renderer"], render);
  const selectedCases = args.case
    ? definitions.cases.filter(({ id }) => id === args.case)
    : definitions.cases;
  if (!selectedCases.length) throw new Error(`Unknown --case ${args.case}`);

  const { default: JSZip } = await import("jszip");
  const oracle = await import("./pptx-programmable-import-oracle.mjs");
  const renderCache = path.join(runRoot, "render-cache");
  if (render) await mkdir(renderCache);
  const results = [];
  for (const caseDefinition of selectedCases) {
    const source = definitions.sources.find(({ id }) => id === caseDefinition.sourceId);
    if (!source) throw new Error(`${caseDefinition.id}: unknown source ${caseDefinition.sourceId}`);
    const sourcePath = source.kind === "controlled"
      ? path.resolve(repoRoot, source.repositoryPath)
      : path.join(assetsDir, source.fileName);
    const sourceBytes = await readFile(sourcePath);
    if (sha256(sourceBytes) !== source.sha256) throw new Error(`${source.id}: source SHA-256 mismatch`);
    const runs = [];
    for (let repetition = 1; repetition <= repetitions; repetition += 1) {
      const runDir = path.join(runRoot, "runs", caseDefinition.id, String(repetition));
      await mkdir(runDir, { recursive: true });
      const inputPath = path.join(runDir, "source.pptx");
      const baselinePath = path.join(runDir, "clone-baseline.pptx");
      const outputPath = path.join(runDir, "output.pptx");
      const receiptPath = path.join(runDir, "worker.json");
      await copyFile(sourcePath, inputPath, constants.COPYFILE_EXCL);
      await chmod(inputPath, 0o444);
      const child = spawnSync(process.execPath, [
        officekitBin,
        "run",
        fileURLToPath(import.meta.url),
        "--worker",
        "--definitions", definitionsPath,
        "--case", caseDefinition.id,
        "--input", inputPath,
        "--baseline", baselinePath,
        "--output", outputPath,
        "--receipt", receiptPath,
      ], { cwd: runDir, encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
      await writeFile(path.join(runDir, "worker.stdout.txt"), child.stdout || "", { flag: "wx" });
      await writeFile(path.join(runDir, "worker.stderr.txt"), child.stderr || "", { flag: "wx" });
      let record;
      try {
        if (child.status !== 0) throw new Error(`public worker exited ${child.status}: ${(child.stderr || child.stdout || "").trim()}`);
        const [baselineBytes, outputBytes, receipt] = await Promise.all([
          readFile(baselinePath),
          readFile(outputPath),
          readFile(receiptPath, "utf8").then(JSON.parse),
        ]);
        const sourceAfter = sha256(await readFile(inputPath));
        if (sourceAfter !== source.sha256) throw new Error(`source copy changed: ${sourceAfter}`);
        const packageOracle = await evaluateCompanionPackageOracle({ JSZip, baselineBytes, outputBytes, caseDefinition, receipt });
        const pixelOracle = render
          ? await evaluateCompanionPixelOracle({ oracle, baselinePath, outputPath, baselineBytes, outputBytes, caseDefinition, receipt, renderCache, targetRenderer })
          : { passed: false, skipped: true, reason: "--no-render" };
        record = {
          repetition,
          status: "passed",
          sourceSha256After: sourceAfter,
          baselineSha256: sha256(baselineBytes),
          outputSha256: sha256(outputBytes),
          canonicalOpcSha256: await canonicalOpcSha256(JSZip, outputBytes),
          worker: receipt,
          packageOracle,
          pixelOracle,
        };
      } catch (error) {
        record = { repetition, status: "failed", reason: errorMessage(error) };
      }
      runs.push(record);
    }
    const passing = runs.filter(({ status }) => status === "passed");
    const deterministic = passing.length === repetitions &&
      new Set(passing.map(({ outputSha256 }) => outputSha256)).size === 1 &&
      new Set(passing.map(({ canonicalOpcSha256 }) => canonicalOpcSha256)).size === 1 &&
      new Set(passing.map(({ packageOracle }) => sha256(Buffer.from(JSON.stringify(packageOracle))))).size === 1;
    results.push({
      id: caseDefinition.id,
      sourceId: source.id,
      covers: caseDefinition.covers,
      requiredRuns: repetitions,
      completedRuns: runs.length,
      passedRuns: passing.length,
      deterministic,
      runs,
    });
  }

  const existingEvidence = await evaluateExistingEvidence(definitions.existingEvidence || []);
  const covered = new Set([
    ...results.filter(({ deterministic }) => deterministic).flatMap(({ covers }) => covers),
    ...existingEvidence.filter(({ passed }) => passed).flatMap(({ covers }) => covers),
  ]);
  const missingCoverage = definitions.requiredCoverage.filter((item) => !covered.has(item));
  const selectedAcceptancePassed = results.every(({ deterministic }) => deterministic);
  const fullAcceptancePassed = selectedAcceptancePassed && missingCoverage.length === 0;
  const evidence = {
    schema: "office-kit/pptx-source-derived-companion-evidence/v1",
    productBaseline: definitions.productBaseline,
    definitionsSha256: sha256(definitionsBytes),
    package: {
      name: packageMetadata.name,
      version: packageMetadata.version,
      installKind: args["install-kind"] || (packageRoot === repoRoot ? "repository" : "packed-clean-install"),
      tarballSha256: args["tarball-sha256"] || null,
    },
    environment: {
      platform: process.platform,
      arch: process.arch,
      node: process.version,
      render,
      targetRenderer,
    },
    repetitionsPerCase: repetitions,
    cases: results,
    existingEvidence,
    notApplicable: definitions.notApplicable,
    acceptance: {
      scope: args.case ? "selected-case" : "full-suite",
      status: (args.case ? selectedAcceptancePassed : fullAcceptancePassed) ? "passed" : "failed",
    },
    coverage: {
      required: definitions.requiredCoverage,
      passed: [...covered].sort(),
      missing: missingCoverage,
      status: fullAcceptancePassed ? "passed" : (args.case && selectedAcceptancePassed ? "partial" : "failed"),
    },
  };
  const evidencePath = path.join(runRoot, "evidence.json");
  await writeFile(evidencePath, `${JSON.stringify(evidence, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ evidence: evidencePath, coverage: evidence.coverage }, null, 2)}\n`);
  if (evidence.acceptance.status !== "passed") process.exitCode = 1;
}

async function runWorker(args) {
  const definitions = JSON.parse(await readFile(path.resolve(required(args, "definitions")), "utf8"));
  validateDefinitions(definitions);
  const caseDefinition = definitions.cases.find(({ id }) => id === required(args, "case"));
  if (!caseDefinition) throw new Error(`Unknown --case ${args.case}`);
  const source = definitions.sources.find(({ id }) => id === caseDefinition.sourceId);
  const inputPath = path.resolve(required(args, "input"));
  const baselinePath = path.resolve(required(args, "baseline"));
  const outputPath = path.resolve(required(args, "output"));
  const receiptPath = path.resolve(required(args, "receipt"));
  await Promise.all([
    requireAbsent(baselinePath, "clone baseline"),
    requireAbsent(outputPath, "output"),
    requireAbsent(receiptPath, "receipt"),
  ]);
  const { FileBlob, PresentationFile } = await import("office-kit");
  const sourceBlob = await FileBlob.load(inputPath);
  if (sha256(sourceBlob.bytes) !== source.sha256) throw new Error(`${source.id}: worker source SHA-256 mismatch`);
  let presentation = await PresentationFile.importPptx(sourceBlob);
  if (presentation.slides.count !== source.slideCount) throw new Error(`${source.id}: expected ${source.slideCount} source slides`);

  let targetPage;
  if (caseDefinition.operation.kind === "component-text") {
    const records = presentation.inspect({ includeComponentCandidates: true, maxChars: Infinity }).ndjson
      .split("\n").filter(Boolean).map(JSON.parse).filter((record) => record.kind === "componentCandidate");
    const candidate = records.find(({ candidateId }) => candidateId === caseDefinition.operation.candidateId);
    if (!candidate) throw new Error(`${caseDefinition.id}: component candidate was not reissued`);
    presentation.reuseSourceComponent({
      candidateId: candidate.candidateId,
      occurrenceIndex: caseDefinition.operation.occurrenceIndex,
      expectedCandidate: candidate,
    });
    targetPage = source.slideCount + 1;
  } else {
    const sourceSlide = presentation.slides.items[caseDefinition.sourceSlideOrdinal - 1];
    if (!sourceSlide?.cloneCapability?.supported) throw new Error(`${caseDefinition.id}: source slide clone is unsupported`);
    sourceSlide.duplicate().moveTo(source.slideCount);
    targetPage = source.slideCount + 1;
  }
  const cloneBaseline = await PresentationFile.exportPptx(presentation);
  await cloneBaseline.save(baselinePath);
  presentation = await PresentationFile.importPptx(cloneBaseline);
  const targetSlide = presentation.slides.items[targetPage - 1];
  if (!targetSlide) throw new Error(`${caseDefinition.id}: clone target page is missing`);
  const operation = await applyWorkerOperation({ presentation, targetSlide, targetPage, caseDefinition });
  const output = await PresentationFile.exportPptx(presentation);
  await output.save(outputPath);
  const reopened = await PresentationFile.importPptx(output);
  const secondImport = verifyWorkerOperation({ presentation: reopened, targetPage, caseDefinition, operation });
  const receipt = {
    schema: "office-kit/pptx-source-derived-companion-worker/v1",
    caseId: caseDefinition.id,
    sourceId: source.id,
    sourceSha256: source.sha256,
    sourceUnchanged: sha256((await FileBlob.load(inputPath)).bytes) === source.sha256,
    sourceSlideCount: source.slideCount,
    targetPage,
    publicApi: operation.publicApi,
    operation,
    baselineSha256: sha256(cloneBaseline.bytes),
    outputSha256: sha256(output.bytes),
    secondImport,
  };
  await writeFile(receiptPath, `${JSON.stringify(receipt, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ caseId: caseDefinition.id, outputSha256: receipt.outputSha256 })}\n`);
}

async function applyWorkerOperation({ presentation, targetSlide, targetPage, caseDefinition }) {
  const operation = caseDefinition.operation;
  if (operation.kind === "native-leaf") {
    const targetId = `presentation/slide/${targetPage}/element/${operation.targetElementOrdinal}`;
    const records = presentation.inspect({ includeNativeLeaves: true, target: targetId, maxChars: Infinity }).ndjson
      .split("\n").filter(Boolean).map(JSON.parse);
    const matches = records.filter((record) => record.kind === "nativeLeaf" &&
      record.targetId === targetId && record.leafKind === operation.leafKind &&
      sameOptional(record.textLeafIndex, operation.textLeafIndex) &&
      sameOptional(record.seriesIndex, operation.seriesIndex) &&
      sameOptional(record.pointIndex, operation.pointIndex) && record.value === operation.before);
    if (matches.length !== 1) throw new Error(`${caseDefinition.id}: expected one issued native leaf, observed ${matches.length}`);
    const record = matches[0];
    const result = presentation.editNativeLeaf(record.targetId, record.leafId, {
      expectedHash: record.expectedHash,
      value: operation.after,
    });
    return {
      publicApi: "presentation.editNativeLeaf",
      targetId,
      leafKind: operation.leafKind,
      before: operation.before,
      after: operation.after,
      receiptKind: result.leafKind,
    };
  }
  if (operation.kind === "table-cell") {
    const table = targetSlide.tables.items.find(({ name }) => name === operation.name);
    if (!table) throw new Error(`${caseDefinition.id}: table ${operation.name} was not found`);
    const cell = table.getCell(operation.row, operation.column);
    if (cell.value !== operation.before) throw new Error(`${caseDefinition.id}: stale table value ${cell.value}`);
    cell.value = operation.after;
    return {
      publicApi: "table.getCell().value",
      targetId: table.id,
      name: table.name,
      row: operation.row,
      column: operation.column,
      before: operation.before,
      after: operation.after,
    };
  }
  if (operation.kind === "replace-image") {
    const image = targetSlide.images.items.find(({ name }) => name === operation.name);
    if (!image) throw new Error(`${caseDefinition.id}: image ${operation.name} was not found`);
    const beforeDataSha256 = dataUrlSha256(image.dataUrl);
    image.replace({ dataUrl: REPLACEMENT_PNG, fit: operation.fit, crop: operation.crop });
    return {
      publicApi: "image.replace",
      targetId: image.id,
      name: image.name,
      beforeDataSha256,
      afterDataSha256: dataUrlSha256(REPLACEMENT_PNG),
      fit: operation.fit,
      crop: operation.crop,
    };
  }
  if (operation.kind === "component-text") {
    const direct = directElements(targetSlide);
    const element = direct.find((candidate) => candidate?.text && typeof candidate.text.value === "string" && candidate.text.value);
    if (!element) throw new Error(`${caseDefinition.id}: reused component exposed no text element`);
    const records = presentation.inspect({ includeNativeLeaves: true, target: element.id, maxChars: Infinity }).ndjson
      .split("\n").filter(Boolean).map(JSON.parse).filter((record) => record.kind === "nativeLeaf" && record.leafKind === "text" && typeof record.value === "string" && record.value);
    if (!records.length) throw new Error(`${caseDefinition.id}: reused component exposed no issued text leaf`);
    const record = records[0];
    const after = `${record.value}${operation.suffix}`;
    presentation.editNativeLeaf(record.targetId, record.leafId, { expectedHash: record.expectedHash, value: after });
    return {
      publicApi: "presentation.reuseSourceComponent+editNativeLeaf",
      targetId: record.targetId,
      leafKind: record.leafKind,
      before: record.value,
      after,
      candidateId: operation.candidateId,
      occurrenceIndex: operation.occurrenceIndex,
    };
  }
  if (operation.kind === "delete-image") {
    const image = targetSlide.images.items[operation.imageIndex];
    if (!image) throw new Error(`${caseDefinition.id}: image index ${operation.imageIndex} was not found`);
    if (image.deletionCapability?.supported !== true) throw new Error(`${caseDefinition.id}: image deletion capability was not issued`);
    const beforeCount = targetSlide.images.items.length;
    const targetId = image.id;
    const name = image.name;
    image.delete();
    return {
      publicApi: "image.delete",
      targetId,
      name,
      beforeCount,
      afterCount: targetSlide.images.items.length,
    };
  }
  if (operation.kind === "slide-reorder") {
    const beforeIndex = targetSlide.index;
    targetSlide.moveTo(operation.destinationIndex);
    return {
      publicApi: "slide.moveTo",
      beforeIndex,
      afterIndex: targetSlide.index,
      destinationIndex: operation.destinationIndex,
    };
  }
  throw new Error(`${caseDefinition.id}: unsupported worker operation ${operation.kind}`);
}

function verifyWorkerOperation({ presentation, targetPage, caseDefinition, operation }) {
  const definition = caseDefinition.operation;
  if (definition.kind === "slide-reorder") {
    if (operation.afterIndex !== definition.destinationIndex) throw new Error(`${caseDefinition.id}: in-memory slide reorder failed`);
    if (presentation.slides.count <= targetPage - 1) throw new Error(`${caseDefinition.id}: reordered presentation lost a slide`);
    return { passed: true, slideCount: presentation.slides.count, destinationIndex: definition.destinationIndex };
  }
  const targetSlide = presentation.slides.items[targetPage - 1];
  if (!targetSlide) throw new Error(`${caseDefinition.id}: target page missing after second import`);
  if (definition.kind === "native-leaf" || definition.kind === "component-text") {
    const records = presentation.inspect({ includeNativeLeaves: true, target: operation.targetId, maxChars: Infinity }).ndjson
      .split("\n").filter(Boolean).map(JSON.parse);
    const found = records.find((record) => record.targetId === operation.targetId && record.leafKind === operation.leafKind && record.value === operation.after);
    if (!found) throw new Error(`${caseDefinition.id}: edited native leaf did not survive second import`);
    return { passed: true, targetId: operation.targetId, observed: found.value };
  }
  if (definition.kind === "table-cell") {
    const table = targetSlide.tables.items.find(({ name }) => name === definition.name);
    const observed = table?.getCell(definition.row, definition.column)?.value;
    if (observed !== definition.after) throw new Error(`${caseDefinition.id}: table cell did not survive second import`);
    return { passed: true, targetId: table.id, observed };
  }
  if (definition.kind === "replace-image") {
    const image = targetSlide.images.items.find(({ name }) => name === definition.name);
    if (!image || dataUrlSha256(image.dataUrl) !== operation.afterDataSha256 || image.fit !== definition.fit || JSON.stringify(image.crop) !== JSON.stringify(definition.crop)) {
      throw new Error(`${caseDefinition.id}: image replacement did not survive second import`);
    }
    return { passed: true, targetId: image.id, observedDataSha256: dataUrlSha256(image.dataUrl), fit: image.fit, crop: image.crop };
  }
  if (definition.kind === "delete-image") {
    if (targetSlide.images.items.length !== operation.afterCount || targetSlide.images.items.some(({ name }) => name === operation.name)) {
      throw new Error(`${caseDefinition.id}: deleted image remained after second import`);
    }
    return { passed: true, observedCount: targetSlide.images.items.length };
  }
  throw new Error(`${caseDefinition.id}: unsupported second-import verifier`);
}

async function evaluateCompanionPackageOracle({ JSZip, baselineBytes, outputBytes, caseDefinition, receipt }) {
  const [baseline, output] = await Promise.all([JSZip.loadAsync(baselineBytes), JSZip.loadAsync(outputBytes)]);
  const diff = await diffPackage(baseline, output);
  const profile = caseDefinition.packageProfile;
  let targetMask;
  if (profile === "clone-slide-token" || profile === "component-slide-token") {
    assertArray(diff.added, [], `${caseDefinition.id}: added parts`);
    assertArray(diff.removed, [], `${caseDefinition.id}: removed parts`);
    if (diff.changed.length !== 1 || !/^ppt\/slides\/slide\d+[.]xml$/u.test(diff.changed[0])) {
      throw new Error(`${caseDefinition.id}: expected one changed SlidePart, observed ${diff.changed.join(", ")}`);
    }
    const [beforeXml, afterXml] = await Promise.all([zipText(baseline, diff.changed[0]), zipText(output, diff.changed[0])]);
    targetMask = maskTokenChange(beforeXml, afterXml, receipt.operation.before, receipt.operation.after);
  } else if (profile === "clone-chart-data") {
    assertArray(diff.added, [], `${caseDefinition.id}: added parts`);
    assertArray(diff.removed, [], `${caseDefinition.id}: removed parts`);
    const chartPart = one(diff.changed.filter((name) => /^ppt\/charts\/chart\d+[.]xml$/u.test(name)), `${caseDefinition.id}: changed ChartPart`);
    const embeddedPart = one(diff.changed.filter((name) => /^ppt\/embeddings\/.+[.]xlsx$/u.test(name)), `${caseDefinition.id}: changed embedded workbook`);
    if (diff.changed.length !== 2) throw new Error(`${caseDefinition.id}: unexpected chart footprint ${diff.changed.join(", ")}`);
    const [chartBefore, chartAfter] = await Promise.all([zipText(baseline, chartPart), zipText(output, chartPart)]);
    const chartMask = scalarSpanMask(chartBefore, chartAfter, receipt.operation.before, receipt.operation.after);
    const [nestedBeforeBytes, nestedAfterBytes] = await Promise.all([zipBytes(baseline, embeddedPart), zipBytes(output, embeddedPart)]);
    const [nestedBefore, nestedAfter] = await Promise.all([JSZip.loadAsync(nestedBeforeBytes), JSZip.loadAsync(nestedAfterBytes)]);
    const nestedDiff = await diffPackage(nestedBefore, nestedAfter);
    assertArray(nestedDiff.added, [], `${caseDefinition.id}: nested added parts`);
    assertArray(nestedDiff.removed, [], `${caseDefinition.id}: nested removed parts`);
    assertArray(nestedDiff.changed, ["xl/worksheets/sheet1.xml"], `${caseDefinition.id}: nested changed parts`);
    const [sheetBefore, sheetAfter] = await Promise.all([zipText(nestedBefore, nestedDiff.changed[0]), zipText(nestedAfter, nestedDiff.changed[0])]);
    const nestedMask = scalarSpanMask(sheetBefore, sheetAfter, receipt.operation.before, receipt.operation.after);
    targetMask = { passed: chartMask.passed && nestedMask.passed, chart: chartMask, nestedWorkbook: nestedMask };
  } else if (profile === "clone-image-copy-on-write") {
    assertArray(diff.removed, [], `${caseDefinition.id}: removed parts`);
    const addedMedia = one(diff.added.filter((name) => /^ppt\/media\/.+$/u.test(name)), `${caseDefinition.id}: added image media`);
    if (diff.added.length !== 1) throw new Error(`${caseDefinition.id}: unexpected added parts ${diff.added.join(", ")}`);
    const slidePart = one(diff.changed.filter((name) => /^ppt\/slides\/slide\d+[.]xml$/u.test(name)), `${caseDefinition.id}: changed SlidePart`);
    const relPart = one(diff.changed.filter((name) => /^ppt\/slides\/_rels\/slide\d+[.]xml[.]rels$/u.test(name)), `${caseDefinition.id}: changed slide relationships`);
    if (diff.changed.length !== 2) throw new Error(`${caseDefinition.id}: unexpected image footprint ${diff.changed.join(", ")}`);
    const [slideBefore, slideAfter, relBefore, relAfter] = await Promise.all([
      zipText(baseline, slidePart), zipText(output, slidePart), zipText(baseline, relPart), zipText(output, relPart),
    ]);
    const picture = maskPictureReplacement(slideBefore, slideAfter, receipt.operation.name);
    const relationships = maskRelationshipReplacement(relBefore, relAfter, addedMedia);
    targetMask = { passed: picture.passed && relationships.passed, picture, relationships };
  } else if (profile === "clone-image-delete") {
    assertArray(diff.added, [], `${caseDefinition.id}: added parts`);
    assertArray(diff.removed, [], `${caseDefinition.id}: removed parts`);
    const slidePart = one(diff.changed.filter((name) => /^ppt\/slides\/slide\d+[.]xml$/u.test(name)), `${caseDefinition.id}: changed SlidePart`);
    const relPart = one(diff.changed.filter((name) => /^ppt\/slides\/_rels\/slide\d+[.]xml[.]rels$/u.test(name)), `${caseDefinition.id}: changed slide relationships`);
    if (diff.changed.length !== 2) throw new Error(`${caseDefinition.id}: unexpected deletion footprint ${diff.changed.join(", ")}`);
    const [slideBefore, slideAfter, relBefore, relAfter] = await Promise.all([
      zipText(baseline, slidePart), zipText(output, slidePart), zipText(baseline, relPart), zipText(output, relPart),
    ]);
    targetMask = maskPictureDeletion(slideBefore, slideAfter, relBefore, relAfter, receipt.operation.name);
  } else if (profile === "presentation-slide-order") {
    assertArray(diff.added, [], `${caseDefinition.id}: added parts`);
    assertArray(diff.removed, [], `${caseDefinition.id}: removed parts`);
    assertArray(diff.changed, ["ppt/presentation.xml"], `${caseDefinition.id}: changed presentation parts`);
    const [beforeXml, afterXml] = await Promise.all([zipText(baseline, diff.changed[0]), zipText(output, diff.changed[0])]);
    targetMask = maskSlideOrder(beforeXml, afterXml, caseDefinition.operation.destinationIndex);
  } else {
    throw new Error(`${caseDefinition.id}: unknown package profile ${profile}`);
  }
  if (targetMask?.passed !== true) throw new Error(`${caseDefinition.id}: target mask failed`);
  return {
    passed: true,
    baselineSha256: sha256(baselineBytes),
    outputSha256: sha256(outputBytes),
    partSet: {
      passed: profile === "clone-image-copy-on-write"
        ? diff.added.length === 1 && diff.removed.length === 0
        : diff.added.length === 0 && diff.removed.length === 0,
      added: diff.added,
      removed: diff.removed,
    },
    changedParts: diff.changed,
    nonTargetPartsByteIdentical: true,
    targetMask,
  };
}

async function evaluateCompanionPixelOracle({ oracle, baselinePath, outputPath, baselineBytes, outputBytes, caseDefinition, receipt, renderCache, targetRenderer }) {
  const baselineRender = await oracle.renderPresentationPages(baselinePath, renderCache, sha256(baselineBytes));
  const outputRender = await oracle.renderPresentationPages(outputPath, renderCache, sha256(outputBytes));
  if (caseDefinition.operation.kind === "slide-reorder") {
    return compareReorderedPages(baselineRender, outputRender, caseDefinition.operation.destinationIndex);
  }
  const primary = oracle.inspectRenderedPages(baselineRender, outputRender, receipt.targetPage);
  if (primary.targetPageChanged) return { ...primary, passed: true, renderer: "libreoffice" };
  if (targetRenderer !== "keynote") throw new Error(`Target rendered page ${receipt.targetPage} did not change`);
  const keynoteCache = path.join(renderCache, "keynote");
  await mkdir(keynoteCache, { recursive: true });
  const baselineKeynote = await oracle.renderKeynotePresentationPages(baselinePath, keynoteCache, sha256(baselineBytes));
  const outputKeynote = await oracle.renderKeynotePresentationPages(outputPath, keynoteCache, sha256(outputBytes));
  const native = oracle.compareRenderedPages(baselineKeynote, outputKeynote, receipt.targetPage);
  return {
    ...native,
    renderer: "keynote",
    primaryRenderer: {
      renderer: "libreoffice",
      targetPageChanged: false,
      nonTargetPagesPixelIdentical: primary.nonTargetPagesPixelIdentical,
    },
  };
}

async function evaluateExistingEvidence(entries) {
  const results = [];
  for (const entry of entries) {
    const filePath = path.resolve(repoRoot, entry.path);
    const bytes = await readFile(filePath);
    const evidence = JSON.parse(bytes);
    if (evidence.schema !== entry.schema) throw new Error(`${entry.path}: schema mismatch`);
    const sources = evidence.sources || [];
    const passed = entry.requiredSourceIds.every((sourceId) => {
      const source = sources.find(({ id }) => id === sourceId);
      return source?.repetitions === entry.requiredRepetitions && source?.deterministic === true &&
        source?.sourceSlideUnchanged === true && source?.verifiedTarget &&
        Array.isArray(source?.unexpectedOverlayChanges) && source.unexpectedOverlayChanges.length === 0;
    });
    results.push({
      covers: entry.covers,
      path: entry.path,
      sha256: sha256(bytes),
      schema: evidence.schema,
      passed,
    });
  }
  return results;
}

function maskTokenChange(beforeXml, afterXml, before, after) {
  const oldToken = xmlEscape(String(before));
  const newToken = xmlEscape(String(after));
  const occurrences = countOccurrences(afterXml, newToken);
  const masked = occurrences === 1 ? afterXml.replace(newToken, oldToken) : "";
  return { passed: occurrences === 1 && masked === beforeXml, oldToken, newToken, occurrences };
}

function scalarSpanMask(beforeXml, afterXml, before, after) {
  const oldToken = `>${xmlEscape(String(before))}<`;
  const newToken = `>${xmlEscape(String(after))}<`;
  const occurrences = countOccurrences(afterXml, newToken);
  const masked = occurrences === 1 ? afterXml.replace(newToken, oldToken) : "";
  return { passed: occurrences === 1 && masked === beforeXml, oldToken, newToken, occurrences };
}

function maskPictureReplacement(beforeXml, afterXml, name) {
  const before = pictureBlock(beforeXml, name);
  const after = pictureBlock(afterXml, name);
  if (!before || !after) return { passed: false, reason: "picture block missing" };
  const beforeEmbed = before.match(/\br:embed="([^"]+)"/u)?.[1];
  const afterEmbed = after.match(/\br:embed="([^"]+)"/u)?.[1];
  if (!beforeEmbed || !afterEmbed || beforeEmbed === afterEmbed) return { passed: false, reason: "picture relationship did not change" };
  const sourceRects = (value) => [...value.matchAll(/<(?:[A-Za-z_][\w.-]*:)?srcRect\b[^>]*\/>/gu)].map((match) => match[0]);
  const beforeRects = sourceRects(before);
  const afterRects = sourceRects(after);
  if (beforeRects.length > 1 || afterRects.length !== 1) return { passed: false, reason: "ambiguous source rectangle" };
  let masked = after.replace(`r:embed="${afterEmbed}"`, `r:embed="${beforeEmbed}"`);
  masked = beforeRects.length === 0
    ? masked.replace(afterRects[0], "")
    : masked.replace(afterRects[0], beforeRects[0]);
  return {
    passed: masked === before && afterXml.replace(after, masked) === beforeXml,
    beforeSha256: sha256(Buffer.from(before)),
    afterSha256: sha256(Buffer.from(after)),
    sourceRelationshipId: beforeEmbed,
    outputRelationshipId: afterEmbed,
    sourceRectangleAdded: beforeRects.length === 0,
  };
}

function maskRelationshipReplacement(beforeXml, afterXml, addedMedia) {
  const before = relationshipElements(beforeXml);
  const after = relationshipElements(afterXml);
  const addedTarget = `../media/${path.posix.basename(addedMedia)}`;
  const next = after.find((entry) => entry.target === addedTarget);
  if (!next) return { passed: false, addedTarget };
  const retainedAfter = after.filter((entry) => entry !== next).map(({ raw }) => raw);
  const retainedBefore = before.map(({ raw }) => raw);
  const passed = after.length === before.length + 1 &&
    JSON.stringify(retainedAfter) === JSON.stringify(retainedBefore) &&
    afterXml.replace(next.raw, "") === beforeXml;
  return { passed, addedTarget, relationshipId: next.id, relationshipType: next.type, sourceRelationshipsRetained: passed };
}

function maskPictureDeletion(beforeSlide, afterSlide, beforeRels, afterRels, name) {
  const block = pictureBlock(beforeSlide, name);
  if (!block || beforeSlide.replace(block, "") !== afterSlide) return { passed: false, reason: "picture block mismatch" };
  const relationshipId = block.match(/\br:embed="([^"]+)"/u)?.[1];
  if (!relationshipId) return { passed: false, reason: "picture relationship missing" };
  const before = relationshipElements(beforeRels);
  const removed = before.find(({ id }) => id === relationshipId);
  if (!removed) return { passed: false, reason: "relationship record missing" };
  const expected = beforeRels.replace(removed.raw, "");
  return { passed: expected === afterRels, relationshipId, removedTarget: removed.target, pictureSha256: sha256(Buffer.from(block)) };
}

function maskSlideOrder(beforeXml, afterXml, destinationIndex) {
  const pattern = /<p:sldIdLst\b[^>]*>([\s\S]*?)<\/p:sldIdLst>/u;
  const beforeMatch = beforeXml.match(pattern);
  const afterMatch = afterXml.match(pattern);
  if (!beforeMatch || !afterMatch) return { passed: false, reason: "slide list missing" };
  const ids = [...beforeMatch[1].matchAll(/<p:sldId\b[^>]*\/>/gu)].map(([raw]) => raw);
  const expected = [...ids];
  const [moved] = expected.splice(expected.length - 1, 1);
  expected.splice(destinationIndex, 0, moved);
  const expectedBody = beforeMatch[1].replace(ids.join(""), expected.join(""));
  const masked = afterXml.replace(afterMatch[1], beforeMatch[1]);
  return { passed: afterMatch[1] === expectedBody && masked === beforeXml, beforeCount: ids.length, destinationIndex };
}

function compareReorderedPages(before, after, destinationIndex) {
  const beforeHashes = before.pages.map(({ sha256: value }) => value);
  const expected = [...beforeHashes];
  const [moved] = expected.splice(expected.length - 1, 1);
  expected.splice(destinationIndex, 0, moved);
  const afterHashes = after.pages.map(({ sha256: value }) => value);
  const passed = JSON.stringify(expected) === JSON.stringify(afterHashes);
  if (!passed) throw new Error("Reordered rendered pages do not match the expected source-page sequence");
  return { passed: true, renderer: "libreoffice", pageContentPixelIdentical: true, destinationIndex, pageCount: afterHashes.length };
}

function pictureBlock(xml, name) {
  const blocks = xml.match(/<p:pic\b[\s\S]*?<\/p:pic>/gu) || [];
  if (name) return blocks.find((block) => block.includes(`name="${xmlEscape(name)}"`));
  return blocks.length === 1 ? blocks[0] : undefined;
}

function relationshipElements(xml) {
  return [...xml.matchAll(/<Relationship\b[^>]*\/>/gu)].map(([raw]) => ({
    raw,
    id: raw.match(/\bId="([^"]+)"/u)?.[1],
    type: raw.match(/\bType="([^"]+)"/u)?.[1],
    target: raw.match(/\bTarget="([^"]+)"/u)?.[1],
  }));
}

async function diffPackage(before, after) {
  const beforeNames = objectPartNames(before);
  const afterNames = objectPartNames(after);
  const beforeSet = new Set(beforeNames);
  const afterSet = new Set(afterNames);
  const added = afterNames.filter((name) => !beforeSet.has(name));
  const removed = beforeNames.filter((name) => !afterSet.has(name));
  const changed = [];
  for (const name of beforeNames.filter((candidate) => afterSet.has(candidate))) {
    const [left, right] = await Promise.all([zipBytes(before, name), zipBytes(after, name)]);
    if (!left.equals(right)) changed.push(name);
  }
  return { added, removed, changed };
}

async function canonicalOpcSha256(JSZip, bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const hash = createHash("sha256");
  for (const name of objectPartNames(zip)) {
    const part = await zipBytes(zip, name);
    hash.update(`${name.length}:${name}:${part.length}:`, "utf8");
    hash.update(part);
  }
  return hash.digest("hex");
}

function objectPartNames(zip) {
  return Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
}

async function zipBytes(zip, name) {
  const file = zip.file(name);
  if (!file) throw new Error(`Missing package part ${name}`);
  return Buffer.from(await file.async("uint8array"));
}

async function zipText(zip, name) {
  const file = zip.file(name);
  if (!file) throw new Error(`Missing package part ${name}`);
  return file.async("text");
}

function directElements(slide) {
  return [
    ...(slide?.shapes?.items || []),
    ...(slide?.tables?.items || []),
    ...(slide?.charts?.items || []),
    ...(slide?.images?.items || []),
    ...(slide?.groups?.items || []),
    ...(slide?.nativeObjects?.items || []),
    ...(slide?.connectors?.items || []),
  ];
}

function dataUrlSha256(value) {
  const match = String(value || "").match(/^data:[^,]*;base64,(.+)$/u);
  if (!match) throw new Error("Expected a base64 data URL");
  return sha256(Buffer.from(match[1], "base64"));
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function xmlEscape(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

function countOccurrences(value, needle) {
  if (!needle) return 0;
  return value.split(needle).length - 1;
}

function one(values, label) {
  if (values.length !== 1) throw new Error(`${label}: expected one, observed ${values.length}`);
  return values[0];
}

function assertArray(actual, expected, label) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) throw new Error(`${label}: expected ${JSON.stringify(expected)}, observed ${JSON.stringify(actual)}`);
}

function sameOptional(actual, expected) {
  return expected == null || actual === expected;
}

function validateDefinitions(value) {
  if (value?.schema !== "office-kit/pptx-source-derived-companion-cases/v1") throw new Error("Unsupported companion definition schema");
  if (!Array.isArray(value.sources) || !Array.isArray(value.cases) || !Array.isArray(value.requiredCoverage)) throw new Error("Invalid companion definitions");
  if (!Number.isInteger(value.repetitionsPerCase) || value.repetitionsPerCase < 1) throw new Error("Invalid companion repetition count");
  if (typeof value.assetsEnvironment !== "string" || !/^[A-Z][A-Z0-9_]*$/u.test(value.assetsEnvironment)) throw new Error("Invalid companion assets environment");
  const sourceIds = new Set(value.sources.map(({ id }) => id));
  for (const item of value.cases) {
    if (!item.id || !sourceIds.has(item.sourceId) || !Array.isArray(item.covers) || !item.operation?.kind || !item.packageProfile) {
      throw new Error(`Invalid companion case ${item?.id || "unknown"}`);
    }
  }
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) throw new Error(`Unexpected argument ${token}`);
    const name = token.slice(2);
    if (new Set(["worker", "no-render"]).has(name)) result[name] = true;
    else if (argv[index + 1] && !argv[index + 1].startsWith("--")) result[name] = argv[++index];
    else throw new Error(`Missing value for ${token}`);
  }
  return result;
}

function parseTargetRenderer(value, render) {
  if (!value) return null;
  if (!render) throw new Error("--target-renderer cannot be used with --no-render");
  if (value !== "keynote") throw new Error(`Unsupported --target-renderer ${value}; expected keynote`);
  if (process.platform !== "darwin") throw new Error("--target-renderer keynote requires macOS");
  return value;
}

function positiveInteger(value, name) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > 10) throw new Error(`${name} must be an integer from 1 through 10`);
  return parsed;
}

function required(args, name) {
  if (!args[name]) throw new Error(`Missing --${name}`);
  return args[name];
}

async function requireAbsent(target, label) {
  try {
    await stat(target);
  } catch (error) {
    if (error.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists; outputs are create-only: ${target}`);
}

function errorMessage(error) {
  return error instanceof Error ? error.message : String(error);
}

await main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
