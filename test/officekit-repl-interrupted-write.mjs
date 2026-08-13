import assert from "node:assert/strict";
import { lstat, mkdtemp, readFile, readdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawn } from "node:child_process";

import { createReplSession } from "../src/cli/repl.mjs";

const repositoryRoot = path.resolve(import.meta.dirname, "..");
const replModule = new URL("../src/cli/repl.mjs", import.meta.url).href;

async function runInterruptedWriteCase(point) {
  const workspaceRoot = await mkdtemp(path.join(os.tmpdir(), `officekit-repl-interrupt-${point}-`));
  const childSource = `
    import { createReplSession } from ${JSON.stringify(replModule)};
    const workspaceRoot = ${JSON.stringify(workspaceRoot)};
    const first = await createReplSession({ workspaceRoot, newTaskGoal: "Interrupted task" });
    await first.handleLine(JSON.stringify({ id: "stable", code: "ctx.state.value = 'process-only'; return ctx.state.value;" }));
    const taskId = first.ready.task.id;
    await first.close();
    process.env.OFFICE_KIT_REPL_TEST_INTERRUPT_AT = ${JSON.stringify(point)};
    const resumed = await createReplSession({ workspaceRoot, taskId });
    await resumed.handleLine(JSON.stringify({ id: "crash", code: "ctx.state.value = 'uncertain'; return ctx.state.value;" }));
  `;
  const child = spawn(process.execPath, ["--input-type=module", "-e", childSource], {
    cwd: repositoryRoot,
    env: { ...process.env, NODE_ENV: "test" },
    stdio: ["ignore", "pipe", "pipe"],
  });
  let stderr = "";
  child.stderr.setEncoding("utf8");
  child.stderr.on("data", (chunk) => { stderr += chunk; });
  const result = await new Promise((resolve, reject) => {
    child.once("error", reject);
    child.once("close", (code, signal) => resolve({ code, signal }));
  });
  assert.equal(result.code, 86, `${point} child did not stop at the injected interruption: ${stderr}`);

  const tasksRoot = path.join(workspaceRoot, ".office-kit", "tasks");
  const taskIds = (await readdir(tasksRoot)).filter((name) => name.startsWith("t_"));
  assert.equal(taskIds.length, 1);
  const taskId = taskIds[0];
  const taskRoot = path.join(tasksRoot, taskId);
  const sessionIds = await readdir(path.join(taskRoot, "sessions"));
  assert.equal(sessionIds.length, 2);
  const sessions = await Promise.all(sessionIds.map(async (sessionId) => {
    const sessionRoot = path.join(taskRoot, "sessions", sessionId);
    const journal = (await readFile(path.join(sessionRoot, "session.jsonl"), "utf8"))
      .trim().split(/\r?\n/u).filter(Boolean).map((line) => JSON.parse(line));
    return { sessionId, sessionRoot, journal };
  }));
  const crashed = sessions.find((session) => session.journal.at(-1)?.id === "crash");
  assert.ok(crashed);
  assert.equal(crashed.journal.at(-1).type, "request.started");
  const checkpointPath = path.join(crashed.sessionRoot, "checkpoint.json");
  const temporaryNames = (await readdir(crashed.sessionRoot)).filter((name) => name.startsWith(".checkpoint.json.") && name.endsWith(".tmp"));
  if (point === "checkpoint-before-rename") {
    await assert.rejects(readFile(checkpointPath), (error) => error.code === "ENOENT");
    assert.equal(temporaryNames.length, 1);
    const temporary = await lstat(path.join(crashed.sessionRoot, temporaryNames[0]));
    assert.equal(temporary.isFile(), true);
    assert.equal(temporary.isSymbolicLink(), false);
  } else {
    const checkpoint = JSON.parse(await readFile(checkpointPath, "utf8"));
    assert.equal(checkpoint.sequence, 1);
    assert.equal(checkpoint.state.safe.value, "uncertain");
    assert.equal(temporaryNames.length, 0);
  }

  const recovered = await createReplSession({ workspaceRoot, taskId });
  assert.equal(recovered.ready.task.state, "attention");
  assert.ok(recovered.ready.task.pending.some((entry) => entry.type === "interrupted-request" && entry.requestId === "crash" && entry.maybeApplied));
  const recovery = await recovered.handleLine(JSON.stringify({
    id: "recovery",
    code: "return {processState: ctx.state.value ?? null, stableHead: ctx.task.head};",
  }));
  assert.equal(recovery.ok, true);
  assert.equal(recovery.audit.maybeApplied, true);
  assert.equal(recovery.audit.interruptedRequest.requestId, "crash");
  assert.equal(recovery.result.processState, null, "a new task session never restores the interrupted JavaScript heap");
  assert.equal(recovery.result.stableHead, null, "an interrupted unreviewed cell never creates a stable commit");
  await recovered.close();
  return { point, platform: process.platform };
}

const results = [];
for (const point of ["checkpoint-before-rename", "checkpoint-after-rename"]) {
  results.push(await runInterruptedWriteCase(point));
}

console.log(`OfficeKit REPL interrupted-write matrix ok (${results.map((item) => `${item.platform}:${item.point}`).join(", ")})`);
