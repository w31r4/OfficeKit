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
  const taskRoot = path.join(workspaceRoot, "task");
  const childSource = `
    import { createReplSession } from ${JSON.stringify(replModule)};
    import { readdir } from "node:fs/promises";
    import path from "node:path";
    const workspaceRoot = ${JSON.stringify(workspaceRoot)};
    const taskRoot = ${JSON.stringify(taskRoot)};
    const first = await createReplSession({ workspaceRoot, taskRoot });
    await first.handleLine(JSON.stringify({ id: "stable", code: "ctx.state.value = 'stable'; return ctx.state.value;" }));
    await first.close();
    const sessionIds = await readdir(path.join(taskRoot, ".officekit-repl"));
    if (sessionIds.length !== 1) throw new Error("expected one checkpoint session");
    const checkpoint = path.join(taskRoot, ".officekit-repl", sessionIds[0], "checkpoint.json");
    process.env.OFFICE_KIT_REPL_TEST_INTERRUPT_AT = ${JSON.stringify(point)};
    const resumed = await createReplSession({ resume: checkpoint });
    await resumed.handleLine(JSON.stringify({ id: "crash", code: "ctx.state.value = 'new'; return ctx.state.value;" }));
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

  const sessionIds = await readdir(path.join(taskRoot, ".officekit-repl"));
  assert.equal(sessionIds.length, 1);
  const sessionRoot = path.join(taskRoot, ".officekit-repl", sessionIds[0]);
  const checkpointPath = path.join(sessionRoot, "checkpoint.json");
  const checkpoint = JSON.parse(await readFile(checkpointPath, "utf8"));
  const journal = (await readFile(path.join(sessionRoot, "session.jsonl"), "utf8"))
    .trim()
    .split(/\r?\n/u)
    .map((line) => JSON.parse(line));
  const temporaryNames = (await readdir(sessionRoot))
    .filter((name) => name.startsWith(".checkpoint.json.") && name.endsWith(".tmp"));
  if (point === "checkpoint-before-rename") {
    assert.equal(checkpoint.sequence, 1, "the previous checkpoint must remain authoritative");
    assert.equal(checkpoint.state.safe.value, "stable");
    assert.equal(temporaryNames.length, 1, "an abrupt pre-rename stop leaves only an isolated temp file");
    const temporary = await lstat(path.join(sessionRoot, temporaryNames[0]));
    assert.equal(temporary.isFile(), true);
    assert.equal(temporary.isSymbolicLink(), false);
  } else {
    assert.equal(checkpoint.sequence, 2, "the renamed checkpoint must be durable before journal terminal append");
    assert.equal(checkpoint.state.safe.value, "new");
    assert.equal(temporaryNames.length, 0);
  }
  assert.equal(journal.at(-1).type, "request.started");
  assert.equal(journal.at(-1).id, "crash");

  const resumed = await createReplSession({ resume: checkpointPath });
  const recovery = await resumed.handleLine(JSON.stringify({
    id: "recovery",
    code: "return {value: ctx.state.value, interrupted: ctx.state.value === 'stable' || ctx.state.value === 'new'};",
  }));
  assert.equal(recovery.ok, true);
  assert.equal(recovery.audit.maybeApplied, true);
  assert.equal(recovery.audit.interruptedRequest.id, "crash");
  assert.equal(recovery.result.interrupted, true);
  await resumed.close();

  const finalCheckpoint = JSON.parse(await readFile(checkpointPath, "utf8"));
  assert.equal(finalCheckpoint.last.id, "recovery");
  const finalJournal = (await readFile(path.join(sessionRoot, "session.jsonl"), "utf8"))
    .trim()
    .split(/\r?\n/u)
    .map((line) => JSON.parse(line));
  assert.equal(finalJournal.at(-1).type, "request.terminal");
  assert.equal(finalJournal.at(-1).id, "recovery");
  return { point, platform: process.platform, checkpointSequence: checkpoint.sequence };
}

const results = [];
for (const point of ["checkpoint-before-rename", "checkpoint-after-rename"]) {
  results.push(await runInterruptedWriteCase(point));
}

console.log(`OfficeKit REPL interrupted-write matrix ok (${results.map((item) => `${item.platform}:${item.point}`).join(", ")})`);
