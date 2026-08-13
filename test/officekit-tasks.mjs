import assert from "node:assert/strict";
import { mkdir, mkdtemp, readFile, symlink, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  acquireTaskLock,
  createTask,
  deleteTask,
  listTasks,
  openTask,
  resolveTaskWorkspace,
  taskDetail,
} from "../src/cli/task-store.mjs";
import { formatTaskList } from "../src/cli/tasks.mjs";

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
