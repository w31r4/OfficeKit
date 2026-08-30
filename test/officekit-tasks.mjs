import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, symlink, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  acquireTaskLock,
  createTask,
  deleteTask,
  listTasks,
  openTask,
  recordTaskPpjRevision,
  resolveTaskWorkspace,
  resumeTaskPpjRevision,
  taskDetail,
} from "../src/cli/task-store.mjs";
import { formatTaskList, runTasksCommand } from "../src/cli/tasks.mjs";
import { resumePpjTask } from "../src/ppj/cli.mjs";

const root = await mkdtemp(path.join(os.tmpdir(), "officekit-tasks-"));
const workspace = path.join(root, "workspace");
const otherWorkspace = path.join(root, "other");
await mkdir(path.join(workspace, "nested", "deep"), { recursive: true });
await mkdir(otherWorkspace);

const created = [];
for (let index = 0; index < 8; index += 1) {
  created.push(await createTask({
    workspaceRoot: workspace,
    goal: `Task ${index + 1}`,
    now: new Date(Date.UTC(2026, 7, 13, 8, index)),
  }));
}
const foreign = await createTask({ workspaceRoot: otherWorkspace, goal: "Foreign task" });

let createdByCli = "";
await runTasksCommand([
  "--new",
  "PPJ import continuation",
  "--workspace",
  otherWorkspace,
  "--json",
], { output: { write(value) { createdByCli += value; } } });
const createdByCliReceipt = JSON.parse(createdByCli);
assert.equal(createdByCliReceipt.task.goal, "PPJ import continuation");
assert.equal(createdByCliReceipt.task.state, "new");
assert.match(createdByCliReceipt.task.id, /^t_[a-f0-9]{12}$/u);

assert.equal(
  await resolveTaskWorkspace({ cwd: path.join(workspace, "nested", "deep") }),
  await (await import("node:fs/promises")).realpath(workspace),
  "the nearest .office-kit directory defines the workspace",
);

const recent = await listTasks({ workspaceRoot: workspace });
assert.equal(recent.total, 8);
assert.equal(recent.shown, 5);
assert.equal(recent.truncated, true);
assert.equal(recent.tasks[0].goal, "Task 8");
assert.equal(recent.tasks.at(-1).goal, "Task 4");
assert.ok(recent.tasks.every((task) => task.id !== foreign.manifest.id));
const human = formatTaskList(recent);
assert.match(human, /OfficeKit .* 8 tasks/);
assert.match(human, /3 more tasks/);
assert.doesNotMatch(human, /Task 1/u);

const all = await listTasks({ workspaceRoot: workspace, all: true });
assert.equal(all.shown, 8);
assert.equal(all.truncated, false);

const selectedId = created[3].manifest.id;
const detail = await taskDetail({ workspaceRoot: workspace, taskId: selectedId });
assert.equal(detail.task.goal, "Task 4");
assert.equal(detail.task.state, "new");
assert.equal(detail.task.head, null);
assert.deepEqual(detail.task.inputs, []);

const ppjProgram = Buffer.from(`${JSON.stringify({
  schema: "office-kit/ppj/v1",
  meta: { id: "task-program", title: "Task program", language: "en-US", version: 1 },
  intent: {},
  design: {},
  assets: [],
  components: [],
  pages: [{ id: "page-1", elements: [] }],
})}\n`);
const digest = (value) => createHash("sha256").update(value).digest("hex");
const programSha256 = digest(ppjProgram);
const nodeMap = Buffer.from('{"schema":"office-kit/ppj-node-map/v1","nodes":[]}\n');
const candidate = Buffer.from("fake deterministic PPTX candidate");
const outputSha256 = digest(candidate);
const ppjTask = await openTask({ workspaceRoot: workspace, taskId: selectedId });
const workspaceState = { program: ppjProgram, source: new Uint8Array(), assets: [] };
const baseReceipt = {
  programJson: ppjProgram,
  programSha256,
  nodeMapJson: nodeMap,
  sourceBound: false,
  restoredEmbeddedProgram: false,
  sourceSha256: "",
  expandedElementCount: 0,
  changedParts: [],
  changedNodeIds: [],
  diagnostics: [],
};
await recordTaskPpjRevision(ppjTask, workspaceState, { stage: "checked", receipt: baseReceipt });
await recordTaskPpjRevision(ppjTask, workspaceState, {
  stage: "built",
  receipt: { ...baseReceipt, outputSha256 },
  candidate: { bytes: candidate, outputPath: path.join(workspace, "task-program.pptx") },
});
await recordTaskPpjRevision(ppjTask, workspaceState, {
  stage: "reviewed",
  receipt: { ...baseReceipt, outputSha256 },
  candidate: { bytes: candidate },
  review: {
    verdict: "passed-with-limitations",
    visualReview: "unavailable",
    playbackEvidence: "structural",
    delivery: { sha256: outputSha256 },
  },
});
const reopenedPpjTask = await openTask({ workspaceRoot: workspace, taskId: selectedId });
const resumedProgram = await resumeTaskPpjRevision(reopenedPpjTask);
assert.equal(resumedProgram.status, "reviewed");
assert.equal(resumedProgram.sha256, programSha256);
assert.equal(path.isAbsolute(resumedProgram.path), true);
assert.deepEqual(await readFile(resumedProgram.path), ppjProgram);
const materializedProgramPath = path.join(workspace, "resumed", "task-program.ppj");
const materializedProgram = await resumePpjTask({
  taskId: selectedId,
  outputPath: materializedProgramPath,
}, { cwd: workspace });
assert.equal(materializedProgram.status, "reviewed");
assert.equal(materializedProgram.programSha256, programSha256);
assert.equal(materializedProgram.output, materializedProgramPath);
assert.deepEqual(await readFile(materializedProgramPath), ppjProgram);

const sourceBytes = Buffer.from("immutable imported PPTX bytes");
const imageBytes = Buffer.from("immutable image bytes");
const sourceProgram = Buffer.from(`${JSON.stringify({
  schema: "office-kit/ppj/v1",
  meta: { id: "source-task", title: "Source task", language: "en-US", version: 1 },
  intent: {},
  design: {},
  source: { uri: "source-assets/source.pptx", sha256: digest(sourceBytes), revision: digest(sourceBytes) },
  assets: [{ id: "asset-image", uri: "source-assets/image.png", mimeType: "image/png", sha256: digest(imageBytes) }],
  components: [],
  pages: [{ id: "page-1", elements: [] }],
})}\n`);
const sourceProgramSha256 = digest(sourceProgram);
const sourceTaskId = created[4].manifest.id;
const sourceTask = await openTask({ workspaceRoot: workspace, taskId: sourceTaskId });
await recordTaskPpjRevision(sourceTask, {
  program: sourceProgram,
  source: sourceBytes,
  assets: [{ id: "asset-image", mimeType: "image/png", data: imageBytes }],
}, {
  stage: "checked",
  receipt: {
    ...baseReceipt,
    programJson: sourceProgram,
    programSha256: sourceProgramSha256,
    sourceBound: true,
    sourceSha256: digest(sourceBytes),
  },
});
const sourceResumePath = path.join(workspace, "resumed-source", "deck.ppj");
const sourceResume = await resumePpjTask({ taskId: sourceTaskId, outputPath: sourceResumePath }, { cwd: workspace });
assert.equal(sourceResume.sourceBound, true);
assert.deepEqual(await readFile(sourceResumePath), sourceProgram);
assert.deepEqual(await readFile(path.join(workspace, "resumed-source", "source-assets", "source.pptx")), sourceBytes);
assert.deepEqual(await readFile(path.join(workspace, "resumed-source", "source-assets", "image.png")), imageBytes);
const ppjDetail = await taskDetail({ workspaceRoot: workspace, taskId: selectedId });
assert.equal(ppjDetail.task.state, "stable");
assert.equal(ppjDetail.task.program.path, resumedProgram.path);

await assert.rejects(
  openTask({ workspaceRoot: workspace, taskId: "defense-deck" }),
  (error) => error.code === "invalid-task-id",
);
await assert.rejects(
  openTask({ workspaceRoot: workspace, taskId: "t_000000000000" }),
  (error) => error.code === "task-not-found",
);

const invalidTask = path.join(workspace, ".office-kit", "tasks", "t_ffffffffffff");
await mkdir(invalidTask);
await writeFile(path.join(invalidTask, "task.json"), "not-json\n");
const withInvalid = await listTasks({ workspaceRoot: workspace, all: true });
assert.equal(withInvalid.total, 8);
assert.ok(withInvalid.invalid.some((entry) => entry.id === "t_ffffffffffff"));

if (process.platform !== "win32") {
  const outside = path.join(root, "outside");
  await mkdir(outside);
  await symlink(outside, path.join(workspace, ".office-kit", "tasks", "t_eeeeeeeeeeee"), "dir");
  const withSymlink = await listTasks({ workspaceRoot: workspace, all: true });
  assert.ok(withSymlink.invalid.some((entry) => entry.id === "t_eeeeeeeeeeee"));
  assert.deepEqual(await readFile(path.join(workspace, ".office-kit", "tasks", ".gitignore"), "utf8"), "*\n!.gitignore\n");
}

const lock = await acquireTaskLock(created[0].taskRoot, { sessionId: "test-session" });
await assert.rejects(
  acquireTaskLock(created[0].taskRoot, { sessionId: "second-session" }),
  (error) => error.code === "task-busy",
);
await assert.rejects(
  deleteTask({ workspaceRoot: workspace, taskId: created[0].manifest.id }),
  (error) => error.code === "task-busy",
);
await lock.release();
const deleted = await deleteTask({ workspaceRoot: workspace, taskId: created[0].manifest.id });
assert.equal(deleted.deleted, true);
assert.ok(deleted.bytes > 0);
assert.equal((await listTasks({ workspaceRoot: workspace, all: true })).total, 7);

console.log("OfficeKit task discovery smoke ok");
