#!/usr/bin/env node

import crypto from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { FileBlob, PresentationFile } from "office-kit";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const definitions = JSON.parse(await readFile(path.resolve(required(args, "definitions")), "utf8"));
  if (definitions?.schema !== "office-kit/pptx-codex-continuation-tasks/v1") throw new Error("Unsupported continuation definitions");
  const task = definitions.tasks.find(({ id }) => id === required(args, "task-id"));
  if (!task) throw new Error(`Unknown continuation task ${args["task-id"]}`);
  const inputPath = path.resolve(required(args, "input"));
  const receiptPath = path.resolve(required(args, "receipt"));
  const bytes = await readFile(inputPath);
  const first = await PresentationFile.importPptx(new FileBlob(bytes, { type: PPTX_MIME, name: path.basename(inputPath) }));
  verifyPresentation(first, task);
  const roundTrip = await PresentationFile.exportPptx(first);
  const second = await PresentationFile.importPptx(new FileBlob(roundTrip.bytes, { type: PPTX_MIME, name: path.basename(inputPath) }));
  const observed = verifyPresentation(second, task);
  const receipt = {
    schema: "office-kit/pptx-programmable-import-continuation-verify/v1",
    taskId: task.id,
    inputSha256: sha256(bytes),
    slideCount: second.slides.count,
    targetPage: task.targetPageAfterAppend,
    observed,
    secondImport: true,
    publicApi: true,
  };
  await writeFile(receiptPath, `${JSON.stringify(receipt, null, 2)}\n`, { flag: "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, taskId: task.id, slideCount: receipt.slideCount })}\n`);
}

function verifyPresentation(presentation, task) {
  if (presentation.slides.count !== task.targetPageAfterAppend) {
    throw new Error(`${task.id}: expected ${task.targetPageAfterAppend} slides, observed ${presentation.slides.count}`);
  }
  const observed = [];
  for (const edit of task.edits) {
    if (edit.operation === "svg-text") {
      const image = presentation.resolve(edit.targetId);
      const matches = (image?.getSvgTextNodes?.() || []).filter((node) => node.id === edit.nodeId && node.text === edit.value);
      if (matches.length !== 1) throw new Error(`${task.id}: expected one ${edit.nodeId} SVG value after second import, observed ${matches.length}`);
      observed.push({ operation: edit.operation, targetId: edit.targetId, nodeId: edit.nodeId, value: matches[0].text });
      continue;
    }
    const records = presentation.inspect({ includeNativeLeaves: true, target: edit.targetId, maxChars: Infinity }).ndjson
      .split("\n").filter(Boolean).map(JSON.parse);
    const matches = records.filter((record) => record.kind === "nativeLeaf" && record.targetId === edit.targetId
      && record.leafKind === edit.leafKind && record.textLeafIndex === edit.textLeafIndex && record.value === edit.value);
    if (matches.length !== 1) throw new Error(`${task.id}: expected one ${edit.leafKind} value after second import, observed ${matches.length}`);
    observed.push({ operation: edit.operation, targetId: edit.targetId, leafId: matches[0].leafId, value: matches[0].value });
  }
  return observed;
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

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

await main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
