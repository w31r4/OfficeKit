import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  buildCodexPrompt,
  evaluateDurableTask,
  inspectAgentRuntimeState,
  scanAgentPolicy,
} from "../scripts/pptx-programmable-import-codex-harness.mjs";
import { sha256 } from "../scripts/pptx-programmable-import-oracle.mjs";

const source = { id: "synthetic", sha256: sha256(Buffer.from("source")), slideCount: 1 };
const task = {
  id: "synthetic-continuation",
  sourceSlide: 1,
  targetPageAfterAppend: 2,
  output: "outputs/result.pptx",
  goal: "Clone and continue the source through three durable sessions.",
  edits: [
    { phase: 1, operation: "native-leaf", targetId: "presentation/slide/2/element/1", leafKind: "text", textLeafIndex: 0, expected: "Before", value: "First" },
    { phase: 2, operation: "native-leaf", targetId: "presentation/slide/2/element/2", leafKind: "text", textLeafIndex: 0, expected: "Before 2", value: "Second" },
  ],
};

const prompt = buildCodexPrompt({ task, source });
assert.match(prompt, /exactly three separate OfficeKit REPL processes/u);
assert.match(prompt, /--file phase-1[.]mjs/u);
assert.match(prompt, /Do not hand-build JSONL/u);
assert.match(prompt, /never launch a second `repl --new`/u);
assert.match(prompt, /empty task-id bootstrap session.*disqualifying/su);
assert.match(prompt, /officekit repl --new/u);
assert.match(prompt, /ctx[.]commit/u);
assert.match(prompt, /ctx[.]publish/u);
assert.match(prompt, /ctx[.]task[.]artifacts[\s\S]*headRevision[\s\S]*path[.]resolve\(ctx[.]taskRoot[\s\S]*FileBlob[.]load/su);
assert.match(prompt, /never call `ctx[.]input\("continued-deck"\)`/u);
assert.match(prompt, /baseline: restoredPath/u);
assert.match(prompt, /application\/octet-stream/u);
assert.match(prompt, /ctx[.]publish\(ctx[.]task[.]commit, \{ artifactId: "continued-deck"/u);
assert.match(prompt, /Presentation[.]create/u);
assert.match(prompt, /inputs\/source[.]pptx/u);

const safePolicy = scanAgentPolicy({
  traceText: `${JSON.stringify({ type: "item.completed", item: { type: "command_execution", command: "node node_modules/office-kit/bin/officekit.mjs repl --new acceptance" } })}\n`,
  authoredFiles: [{ origin: "workspace:task.mjs", text: "const kit = await ctx.import('office-kit'); return kit.PresentationFile;" }],
});
assert.equal(safePolicy.passed, true);
const unsafePolicy = scanAgentPolicy({
  traceText: `${JSON.stringify({ type: "item.completed", item: { type: "command_execution", command: "python3 patch.py && unzip inputs/source.pptx" } })}\n`,
  authoredFiles: [{ origin: "workspace:bad.mjs", text: "Presentation.create(); deck.slides.add({});" }],
});
assert.equal(unsafePolicy.passed, false);
assert.deepEqual(new Set(unsafePolicy.findings.map(({ code }) => code)), new Set(["python", "raw-opc", "whole-rebuild"]));
assert.equal(scanAgentPolicy({ authoredFiles: [{ origin: "repl:probe", text: "Object.getOwnPropertyNames(PresentationFile);" }] }).findings[0].code, "api-reflection");

const workspace = await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-harness-test-"));
const taskId = "t_0123456789ab";
const taskRoot = path.join(workspace, ".office-kit/tasks", taskId);
const outputPath = path.join(workspace, task.output);
await mkdir(path.dirname(outputPath), { recursive: true });
await mkdir(path.join(taskRoot, "revisions/continued-deck"), { recursive: true });
await mkdir(path.join(taskRoot, "evidence"), { recursive: true });
for (const sessionId of ["session-1", "session-2", "session-3"]) await mkdir(path.join(taskRoot, "sessions", sessionId), { recursive: true });
const revision1 = Buffer.from("revision-one");
const revision2 = Buffer.from("revision-two");
const revision1Path = `revisions/continued-deck/${sha256(revision1)}.pptx`;
const revision2Path = `revisions/continued-deck/${sha256(revision2)}.pptx`;
await writeFile(path.join(taskRoot, revision1Path), revision1);
await writeFile(path.join(taskRoot, revision2Path), revision2);
await writeFile(outputPath, revision2);
const commits = [];
for (const [index, revision, revisionPath] of [[1, revision1, revision1Path], [2, revision2, revision2Path]]) {
  const commitId = `c${String(index).padStart(4, "0")}`;
  const evidenceBytes = Buffer.from(`${commitId} evidence\n`);
  const evidencePath = `evidence/${commitId}-continued-deck.json`;
  await writeFile(path.join(taskRoot, evidencePath), evidenceBytes);
  const digest = sha256(revision);
  const review = {
    verdict: "passed-with-limitations",
    deliverySha256: digest,
    evidence: { path: evidencePath, bytes: evidenceBytes.length, sha256: sha256(evidenceBytes) },
  };
  commits.push({
    id: commitId,
    artifactId: "continued-deck",
    revisionSha256: digest,
    review,
    heads: { "continued-deck": { path: revisionPath, sha256: digest, bytes: revision.length, review } },
  });
}
const manifest = {
  schemaVersion: 1,
  id: taskId,
  artifacts: [
    { id: "source-deck", source: { sha256: source.sha256, storedPath: "inputs/source-deck/source.pptx" } },
    { id: "continued-deck", kind: "presentation", source: null },
  ],
  commits,
  head: { commitId: "c0002", artifactId: "continued-deck", revisionSha256: sha256(revision2) },
  pending: [],
  publications: [{ commitId: "c0002", artifactId: "continued-deck", path: outputPath, bytes: revision2.length, sha256: sha256(revision2) }],
  lastSessionId: "session-3",
};
await writeFile(path.join(taskRoot, "task.json"), `${JSON.stringify(manifest, null, 2)}\n`);
const durable = await evaluateDurableTask({ workspace, task, source, outputPath });
assert.equal(durable.passed, true);
assert.equal(durable.sessions, 3);
assert.deepEqual(durable.commits.map(({ commitId }) => commitId), ["c0001", "c0002"]);
assert.equal(await inspectAgentRuntimeState(workspace), null);
await mkdir(path.join(taskRoot, "sessions", "session-4"));
assert.deepEqual(await inspectAgentRuntimeState(workspace), { code: "repl-session-budget-exceeded", observed: 4, maximum: 3 });

manifest.pending.push({ type: "review-failed" });
await writeFile(path.join(taskRoot, "task.json"), `${JSON.stringify(manifest, null, 2)}\n`);
await assert.rejects(evaluateDurableTask({ workspace, task, source, outputPath }), /unresolved pending records/u);
assert.equal((await readFile(outputPath, "utf8")), "revision-two");

console.log("PPTX programmable-import Codex harness contract ok");
