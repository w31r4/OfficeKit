#!/usr/bin/env node

import crypto from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { PresentationFile } from "../src/presentation/index.mjs";
import { FileBlob } from "../src/shared/file-blob.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const definitions = JSON.parse(await readFile(path.resolve(required(args, "definitions")), "utf8"));
  if (definitions?.schema !== "office-kit/pptx-programmable-import-intents/v1") throw new Error("Unsupported intent definitions");
  const { source, intent } = resolveIntent(definitions, required(args, "source-id"), required(args, "intent-id"));
  const inputPath = path.resolve(required(args, "input"));
  const outputPath = path.resolve(required(args, "output"));
  const receiptPath = path.resolve(required(args, "receipt"));
  const sourceBytes = await readFile(inputPath);
  const sourceBefore = sha256(sourceBytes);
  if (sourceBefore !== source.sha256) throw new Error(`${source.id}: source SHA-256 mismatch`);
  const presentation = await PresentationFile.importPptx(new FileBlob(sourceBytes, { type: PPTX_MIME, name: source.fileName }));
  applyIntent(presentation, intent);
  const exported = await PresentationFile.exportPptx(presentation);
  const reopened = await PresentationFile.importPptx(new FileBlob(exported.bytes, { type: PPTX_MIME, name: `${intent.id}.pptx` }));
  verifyIntent(reopened, intent);
  await writeFile(outputPath, exported.bytes, { flag: "wx" });
  const sourceAfter = sha256(await readFile(inputPath));
  if (sourceAfter !== sourceBefore) throw new Error(`${source.id}: input changed during public API execution`);
  const receipt = {
    schema: "office-kit/pptx-programmable-import-worker/v1",
    sourceId: source.id,
    intentId: intent.id,
    sourceSha256: sourceBefore,
    sourceUnchanged: true,
    outputSha256: sha256(exported.bytes),
    outputBytes: exported.bytes.length,
    secondImport: true,
    observedValue: observedValue(reopened, intent),
    publicApi: intent.operation === "svg-text" ? "ImageElement.editSvgText" : "presentation.editNativeLeaf",
    runtimeDiagnostics: exported.metadata?.diagnostics || [],
    runtimeEditPlan: exported.metadata?.editPlan || null,
  };
  await writeFile(receiptPath, `${JSON.stringify(receipt, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, sourceId: source.id, intentId: intent.id, outputSha256: receipt.outputSha256 })}\n`);
}

function applyIntent(presentation, intent) {
  if (intent.operation === "svg-text") {
    const image = presentation.resolve(intent.targetId);
    const nodes = image?.getSvgTextNodes?.() || [];
    const matches = nodes.filter((node) => node.id === intent.nodeId && node.text === intent.expected);
    if (matches.length !== 1) throw new Error(`${intent.id}: expected one SVG text leaf, observed ${matches.length}`);
    image.editSvgText(matches[0].id, { expectedHash: matches[0].expectedHash, value: intent.value });
    return;
  }
  const leaf = resolveNativeLeaf(presentation, intent, intent.expected);
  presentation.editNativeLeaf(leaf.targetId, leaf.leafId, { expectedHash: leaf.expectedHash, value: intent.value });
}

function verifyIntent(presentation, intent) {
  if (intent.operation === "svg-text") {
    const image = presentation.resolve(intent.targetId);
    const matches = (image?.getSvgTextNodes?.() || []).filter((node) => node.id === intent.nodeId && node.text === intent.value);
    if (matches.length !== 1) throw new Error(`${intent.id}: SVG value did not survive second import`);
    return;
  }
  resolveNativeLeaf(presentation, intent, intent.value);
}

function observedValue(presentation, intent) {
  if (intent.operation === "svg-text") return presentation.resolve(intent.targetId).getSvgTextNodes().find((node) => node.id === intent.nodeId)?.text;
  return resolveNativeLeaf(presentation, intent, intent.value).value;
}

function resolveNativeLeaf(presentation, intent, expectedValue) {
  const records = presentation.inspect({ includeNativeLeaves: true, target: intent.targetId, maxChars: Infinity }).ndjson
    .split("\n").filter(Boolean).map((line) => JSON.parse(line));
  const matches = records.filter((record) => record.kind === "nativeLeaf" && record.targetId === intent.targetId && record.leafKind === intent.leafKind
    && (intent.textLeafIndex === undefined || record.textLeafIndex === intent.textLeafIndex)
    && (intent.seriesIndex === undefined || record.seriesIndex === intent.seriesIndex)
    && (intent.pointIndex === undefined || record.pointIndex === intent.pointIndex)
    && record.value === expectedValue);
  if (matches.length !== 1) throw new Error(`${intent.id}: expected one ${intent.leafKind} leaf with ${JSON.stringify(expectedValue)}, observed ${matches.length}`);
  return matches[0];
}

function parseArgs(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--") || !argv[index + 1] || argv[index + 1].startsWith("--")) throw new Error(`Expected --name value, received ${token}`);
    result[token.slice(2)] = argv[++index];
  }
  return result;
}

function required(args, name) {
  if (!args[name]) throw new Error(`Missing --${name}`);
  return args[name];
}

function resolveIntent(definitions, sourceId, intentId) {
  const source = definitions.sources.find((candidate) => candidate.id === sourceId);
  const intent = source?.intents.find((candidate) => candidate.id === intentId);
  if (!source || !intent) throw new Error(`Unknown intent ${sourceId}/${intentId}`);
  return { source, intent };
}

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

await main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
