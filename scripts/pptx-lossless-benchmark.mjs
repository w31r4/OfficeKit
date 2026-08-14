#!/usr/bin/env node

import { createHash } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import JSZip from "jszip";

import { FileBlob, PresentationFile } from "../src/index.mjs";

const MANIFEST_SCHEMA = "office-kit/pptx-lossless-benchmark/v1";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;
const MAX_PARTS = 10_000;
const MAX_PART_BYTES = 128 * 1024 * 1024;
const MAX_TOTAL_PART_BYTES = 2 * 1024 * 1024 * 1024;
const DEFAULT_MANIFEST = path.resolve("evals/pptx-lossless/manifest.v1.json");

const SOURCES = Object.freeze([
  {
    id: "suanzhi-future-2026",
    fileName: "b34ddad8cf8b_012_算秩未来2026_0127_极致技术&长期主义.pptx",
    sha256: "b34ddad8cf8bbd083b60e07f8488267b1a0e4199db422468faa0eeb5d83e1762",
    targets: [
      { id: "title", nodeId: "presentation/slide/1/element/1", expected: "极致技术，长期主义", value: "极致技术，长期验证" },
      { id: "vision", nodeId: "presentation/slide/1/element/2", expected: "Vision and ambition", value: "Vision, proof and ambition" },
      {
        id: "partner",
        nodeId: "presentation/slide/1/element/3",
        expected: "通往AGI顶峰的伙伴",
        search: "顶峰的伙伴",
        value: "顶峰的长期伙伴",
        result: "通往AGI顶峰的长期伙伴",
      },
      {
        id: "group-storage",
        nodeId: "presentation/slide/6/element/8/element/2",
        operation: "nativeLeaf",
        leafKind: "text",
        expectedValue: "快速存储",
        value: "高速存储验证",
      },
    ],
  },
  {
    id: "blue-gray-acid-template",
    fileName: "template.pptx",
    sha256: "558ce85c0d64cd2a06faf88d6a4aa331e8cd4c685c59101c835ded2fbc87696d",
    targets: [
      { id: "subtitle", nodeId: "presentation/slide/1/element/6", expected: "单击此处添加副标题内容", value: "OfficeKit 保真编辑评测" },
      {
        id: "cover-title",
        nodeId: "presentation/slide/1/element/7",
        expected: "蓝灰酸性\n季度工作总结",
        search: "季度工作",
        value: "季度验证",
        result: "蓝灰酸性\n季度验证总结",
      },
      { id: "presenter", nodeId: "presentation/slide/1/element/8", expected: "汇报人：稻小壳", value: "汇报人：OfficeKit" },
    ],
  },
  {
    id: "mckinsey-customer-loyalty",
    fileName: "ppt169_麦肯锡风_kimsoong_customer_loyalty.pptx",
    sha256: "e0bfb89454f51c400ac03797c255aa93919328ff8dba36fe414e5bcfed0536c5",
    targets: [
      {
        id: "image-left",
        nodeId: "presentation/slide/1/element/1",
        operation: "nativeLeaf",
        leafKind: "leftEmu",
        expectedValue: 0,
        value: 9_525,
      },
    ],
  },
]);

const [command = "verify", ...argv] = process.argv.slice(2);
const options = parseArgs(argv);
const assetsDir = path.resolve(options.assetsDir || process.env.OFFICEKIT_PPTX_BENCHMARK_ASSETS || "");
const manifestPath = path.resolve(options.manifest || DEFAULT_MANIFEST);

if (!options.assetsDir && !process.env.OFFICEKIT_PPTX_BENCHMARK_ASSETS) {
  fail("Use --assets-dir <dir> or OFFICEKIT_PPTX_BENCHMARK_ASSETS for the external PPTX sources.");
}

if (command === "freeze") {
  const manifest = await freezeManifest(assetsDir);
  await mkdir(path.dirname(manifestPath), { recursive: true });
  await writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, { flag: options.force ? "w" : "wx" });
  print({ ok: true, command, manifest: manifestPath, sources: manifest.sources.length });
} else if (command === "verify") {
  const manifest = await readManifest(manifestPath);
  const actual = await freezeManifest(assetsDir);
  assertJsonEqual(actual, manifest, "Benchmark manifest drifted from the immutable sources");
  print({ ok: true, command, manifest: manifestPath, sources: manifest.sources.length });
} else if (command === "run") {
  const manifest = await readManifest(manifestPath);
  const outputDir = path.resolve(options.outputDir || "tmp/pptx-lossless-benchmark");
  assertSeparateRoot(assetsDir, outputDir);
  await mkdir(outputDir, { recursive: true });
  const repetitions = positiveInteger(options.repetitions ?? 3, "repetitions", 1, 10);
  if (options.target && !options.source) fail("--target requires --source so the benchmark selection is unambiguous.");
  const evidence = await runBenchmark(manifest, assetsDir, outputDir, repetitions, {
    sourceId: options.source,
    targetId: options.target,
  });
  const evidencePath = path.join(outputDir, "evidence.json");
  await writeFile(evidencePath, `${JSON.stringify(evidence, null, 2)}\n`, { flag: options.force ? "w" : "wx" });
  print({ ok: true, command, evidence: evidencePath, sources: evidence.sources.length, repetitions });
} else {
  fail("Usage: pptx-lossless-benchmark.mjs freeze|verify|run --assets-dir <dir> [--manifest <path>] [--output-dir <dir>] [--repetitions 3] [--source <id> [--target <id>]] [--force]");
}

async function freezeManifest(root) {
  const sources = [];
  for (const definition of SOURCES) {
    const sourcePath = path.join(root, definition.fileName);
    const bytes = await boundedRead(sourcePath);
    const digest = sha256(bytes);
    if (digest !== definition.sha256) fail(`Source hash mismatch for ${definition.fileName}: ${digest}`);
    const inventory = await packageInventory(bytes);
    const nodes = await editableNodeIndex(bytes);
    for (const target of definition.targets) await proveDeclaredTarget(bytes, nodes, definition.id, target);
    sources.push({
      id: definition.id,
      fileName: definition.fileName,
      bytes: bytes.byteLength,
      sha256: digest,
      inventory,
      editableNodes: nodes,
      targets: definition.targets,
    });
  }
  return { schema: MANIFEST_SCHEMA, sources };
}

async function runBenchmark(manifest, root, outputRoot, repetitions, selection = {}) {
  const sources = [];
  const selectedSources = selection.sourceId
    ? manifest.sources.filter((source) => source.id === selection.sourceId)
    : manifest.sources;
  if (!selectedSources.length) fail(`Unknown benchmark source: ${selection.sourceId}`);
  for (const sourceManifest of selectedSources) {
    progress({ phase: "source", source: sourceManifest.id });
    const sourcePath = path.join(root, sourceManifest.fileName);
    const sourceBytes = await boundedRead(sourcePath);
    if (sha256(sourceBytes) !== sourceManifest.sha256) fail(`Source hash mismatch for ${sourceManifest.fileName}.`);
    const noOp = await PresentationFile.exportPptx(await importPresentation(sourceBytes));
    if (!Buffer.from(noOp.bytes).equals(sourceBytes)) fail(`No-op export changed ${sourceManifest.id}.`);
    progress({ phase: "no-op", source: sourceManifest.id, ok: true });
    const edits = [];
    const selectedTargets = selection.targetId
      ? sourceManifest.targets.filter((target) => target.id === selection.targetId)
      : sourceManifest.targets;
    if (selection.targetId && !selectedTargets.length) fail(`Unknown benchmark target: ${sourceManifest.id}/${selection.targetId}`);
    for (const target of selectedTargets) {
      const runs = [];
      for (let repetition = 1; repetition <= repetitions; repetition += 1) {
        progress({ phase: "edit", source: sourceManifest.id, target: target.id, repetition });
        const presentation = await importPresentation(sourceBytes);
        applyBenchmarkTarget(presentation, sourceManifest.id, target);
        const output = await PresentationFile.exportPptx(presentation);
        if (output.metadata.editPlan?.schema !== "office-kit/pptx-edit-plan/v1") fail(`Target ${sourceManifest.id}/${target.id} did not compile to an Edit Plan.`);
        const reopened = await importPresentation(output.bytes);
        verifyBenchmarkTarget(reopened, sourceManifest.id, target);
        const oracle = await packageOracle(sourceBytes, output.bytes, output.metadata.editPlan, target);
        const outputName = `${sourceManifest.id}-${target.id}-${repetition}.pptx`;
        await writeFile(path.join(outputRoot, outputName), output.bytes, { flag: options.force ? "w" : "wx" });
        runs.push({ repetition, output: outputName, sha256: sha256(output.bytes), editPlan: output.metadata.editPlan, oracle });
      }
      const hashes = new Set(runs.map((run) => run.sha256));
      const footprints = new Set(runs.map((run) => sha256(Buffer.from(JSON.stringify(run.editPlan)))));
      if (hashes.size !== 1 || footprints.size !== 1) fail(`Target ${sourceManifest.id}/${target.id} is not deterministic across ${repetitions} clean runs.`);
      edits.push({ targetId: target.id, deterministic: true, runs });
    }
    sources.push({ id: sourceManifest.id, sourceSha256: sourceManifest.sha256, noOpByteIdentical: true, edits });
  }
  return { schema: "office-kit/pptx-lossless-evidence/v1", manifestSha256: sha256(await readFile(manifestPath)), repetitions, sources };
}

async function packageOracle(sourceBytes, outputBytes, editPlan, target) {
  const source = await zipParts(sourceBytes);
  const output = await zipParts(outputBytes);
  if (source.size !== output.size || [...source.keys()].some((partPath) => !output.has(partPath))) fail("Edit changed the OPC entry set.");
  const changedParts = [...source.keys()].filter((partPath) => !source.get(partPath).equals(output.get(partPath))).sort();
  const declared = [...editPlan.changedParts].sort();
  assertJsonEqual(changedParts, declared, "Mutation footprint does not match the actual changed OPC parts");
  for (const partPath of source.keys()) {
    if (!changedParts.includes(partPath) && !source.get(partPath).equals(output.get(partPath))) fail(`Non-target OPC part changed: ${partPath}`);
  }
  for (const partPath of changedParts) {
    const sourceXml = source.get(partPath).toString("utf8");
    const outputXml = output.get(partPath).toString("utf8");
    const operation = editPlan.operations.find((candidate) => candidate.slidePartPath === partPath);
    if (!operation) fail(`Edit Plan has no operation for changed part ${partPath}.`);
    const encodedOld = escapeXmlText(operation.expectedValue);
    const encodedNew = escapeXmlText(operation.value);
    if (singleTokenMaskMatches(sourceXml, outputXml, encodedOld, encodedNew) !== 1) fail(`Declared token cannot uniquely mask target XML back to source: ${partPath}`);
  }
  return { changedParts, nonTargetPartsByteIdentical: true, maskedTargetXmlByteIdentical: true };
}

async function packageInventory(bytes) {
  const parts = await zipParts(bytes);
  const partHashes = Object.fromEntries([...parts.entries()].map(([partPath, value]) => [partPath, sha256(value)]));
  const xml = [...parts.entries()].filter(([partPath]) => /[.]xml$/iu.test(partPath)).map(([, value]) => value.toString("utf8")).join("\n");
  const paths = [...parts.keys()];
  const relationshipCount = [...parts.entries()].filter(([partPath]) => /[.]rels$/iu.test(partPath))
    .reduce((count, [, value]) => count + (value.toString("utf8").match(/<Relationship\b/gu) || []).length, 0);
  return {
    partCount: parts.size,
    relationshipCount,
    slideCount: countPaths(paths, /^ppt\/slides\/slide\d+[.]xml$/iu),
    masterCount: countPaths(paths, /^ppt\/slideMasters\/slideMaster\d+[.]xml$/iu),
    layoutCount: countPaths(paths, /^ppt\/slideLayouts\/slideLayout\d+[.]xml$/iu),
    themeCount: countPaths(paths, /^ppt\/theme\/theme\d+[.]xml$/iu),
    chartCount: countPaths(paths, /^ppt\/charts\/chart\d+[.]xml$/iu),
    mediaCount: countPaths(paths, /^ppt\/media\//iu),
    embeddingCount: countPaths(paths, /^ppt\/embeddings\//iu),
    notesCount: countPaths(paths, /^ppt\/notesSlides\/notesSlide\d+[.]xml$/iu),
    commentPartCount: countPaths(paths, /^ppt\/(?:comments|modernComments)\//iu),
    diagramPartCount: countPaths(paths, /^ppt\/diagrams\//iu),
    groupCount: occurrences(xml, "<p:grpSp>"),
    graphicFrameCount: occurrences(xml, "<p:graphicFrame>"),
    smartArtReferenceCount: occurrences(xml, ":relIds"),
    oleObjectCount: occurrences(xml, "<p:oleObj"),
    timingCount: occurrences(xml, "<p:timing"),
    transitionCount: occurrences(xml, "<p:transition"),
    partHashes,
  };
}

async function editableNodeIndex(bytes) {
  const presentation = await importPresentation(bytes);
  const nodes = [];
  for (const slide of presentation.slides.items) collectNodes(slide, slide.index + 1, nodes, 0);
  return nodes;
}

async function proveDeclaredTarget(bytes, nodes, sourceId, target) {
  if ((target.operation ?? "text") === "text") {
    const match = nodes.find((node) => node.id === target.nodeId);
    if (!match || match.text !== target.expected) fail(`Declared target ${sourceId}/${target.id} does not match the imported node index.`);
    return;
  }
  if (target.operation !== "nativeLeaf") fail(`Declared target ${sourceId}/${target.id} uses an unknown operation.`);
  const presentation = await importPresentation(bytes);
  const leaf = nativeLeafRecord(presentation, target);
  if (leaf.value !== target.expectedValue) fail(`Declared native target ${sourceId}/${target.id} does not match its inspected leaf.`);
}

function applyBenchmarkTarget(presentation, sourceId, target) {
  if ((target.operation ?? "text") === "text") {
    const node = presentation.resolve(target.nodeId);
    if (!node?.text || node.text.value !== target.expected) fail(`Target ${sourceId}/${target.id} is stale.`);
    const search = target.search ?? target.expected;
    const result = target.result ?? target.value;
    node.text.replace(search, target.value);
    if (node.text.value !== result) fail(`Target ${sourceId}/${target.id} did not produce its declared model result.`);
    return;
  }
  const leaf = nativeLeafRecord(presentation, target);
  if (leaf.value !== target.expectedValue) fail(`Target ${sourceId}/${target.id} is stale.`);
  presentation.editNativeLeaf(leaf.targetId, leaf.leafId, { expectedHash: leaf.expectedHash, value: target.value });
}

function verifyBenchmarkTarget(presentation, sourceId, target) {
  if ((target.operation ?? "text") === "text") {
    const result = target.result ?? target.value;
    if (presentation.resolve(target.nodeId)?.text?.value !== result) fail(`Target ${sourceId}/${target.id} failed second import.`);
    return;
  }
  if (nativeLeafRecord(presentation, target).value !== target.value) fail(`Target ${sourceId}/${target.id} failed second import.`);
}

function nativeLeafRecord(presentation, target) {
  const records = presentation.inspect({ includeNativeLeaves: true, target: target.nodeId }).ndjson
    .split("\n")
    .filter(Boolean)
    .map((line) => JSON.parse(line));
  const leaves = records.filter((record) => record.kind === "nativeLeaf" && record.targetId === target.nodeId && record.leafKind === target.leafKind);
  if (leaves.length !== 1) fail(`Native target ${target.id} resolved ${leaves.length} ${target.leafKind} leaves.`);
  return leaves[0];
}

function collectNodes(container, slide, output, depth) {
  for (const shape of container.shapes?.items || []) {
    if (!shape.text?.value) continue;
    output.push({ id: shape.id, slide, depth, kind: "shape", name: shape.name, text: shape.text.value });
  }
  for (const group of container.groups?.items || []) collectNodes(group, slide, output, depth + 1);
}

async function importPresentation(bytes) {
  return PresentationFile.importPptx(new FileBlob(bytes, { type: PPTX_MIME }));
}

async function zipParts(bytes) {
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true });
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  if (names.length > MAX_PARTS) fail(`PPTX has ${names.length} parts and exceeds ${MAX_PARTS}.`);
  let total = 0;
  const parts = new Map();
  for (const name of names) {
    if (path.posix.isAbsolute(name) || name.split("/").includes("..")) fail(`Unsafe OPC part path: ${name}`);
    const value = Buffer.from(await zip.files[name].async("uint8array"));
    if (value.byteLength > MAX_PART_BYTES) fail(`OPC part exceeds ${MAX_PART_BYTES} bytes: ${name}`);
    total += value.byteLength;
    if (total > MAX_TOTAL_PART_BYTES) fail(`PPTX uncompressed content exceeds ${MAX_TOTAL_PART_BYTES} bytes.`);
    parts.set(name, value);
  }
  return parts;
}

async function boundedRead(filePath) {
  const bytes = await readFile(filePath);
  if (bytes.byteLength <= 0 || bytes.byteLength > MAX_SOURCE_BYTES) fail(`PPTX source size is outside 1..${MAX_SOURCE_BYTES}: ${filePath}`);
  return bytes;
}

async function readManifest(filePath) {
  const manifest = JSON.parse(await readFile(filePath, "utf8"));
  if (manifest?.schema !== MANIFEST_SCHEMA || !Array.isArray(manifest.sources) || manifest.sources.length !== SOURCES.length) fail("Benchmark manifest schema is invalid.");
  return manifest;
}

function assertSeparateRoot(sourceRoot, outputRoot) {
  const relative = path.relative(sourceRoot, outputRoot);
  if (relative === "" || (!relative.startsWith(`..${path.sep}`) && relative !== ".." && !path.isAbsolute(relative))) fail("Benchmark output directory must be outside the source asset directory.");
}

function countPaths(paths, pattern) { return paths.filter((value) => pattern.test(value)).length; }
function occurrences(value, token) { return value.split(token).length - 1; }
function singleTokenMaskMatches(source, output, expected, replacement) {
  if (!replacement || expected === replacement) return 0;
  let matches = 0;
  for (let index = output.indexOf(replacement); index >= 0; index = output.indexOf(replacement, index + 1)) {
    if (`${output.slice(0, index)}${expected}${output.slice(index + replacement.length)}` === source) matches += 1;
  }
  return matches;
}
function sha256(value) { return createHash("sha256").update(value).digest("hex"); }
function escapeXmlText(value) { return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;"); }
function assertJsonEqual(actual, expected, message) { if (JSON.stringify(actual) !== JSON.stringify(expected)) fail(message); }
function positiveInteger(value, label, minimum, maximum) { const number = Number(value); if (!Number.isSafeInteger(number) || number < minimum || number > maximum) fail(`${label} must be ${minimum}..${maximum}.`); return number; }
function print(value) { process.stdout.write(`${JSON.stringify(value)}\n`); }
function progress(value) { process.stderr.write(`${JSON.stringify(value)}\n`); }
function fail(message) { const error = new Error(message); error.code = "pptx-lossless-benchmark-failed"; throw error; }

function parseArgs(args) {
  const parsed = {};
  for (let index = 0; index < args.length; index += 1) {
    const token = args[index];
    if (token === "--force") parsed.force = true;
    else if (token === "--assets-dir") parsed.assetsDir = args[++index];
    else if (token === "--manifest") parsed.manifest = args[++index];
    else if (token === "--output-dir") parsed.outputDir = args[++index];
    else if (token === "--repetitions") parsed.repetitions = args[++index];
    else if (token === "--source") parsed.source = args[++index];
    else if (token === "--target") parsed.target = args[++index];
    else fail(`Unknown argument: ${token}`);
  }
  return parsed;
}
