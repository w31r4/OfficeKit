import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { PassThrough } from "node:stream";

import { createReplSession, REPL_PROTOCOL_VERSION, runReplCommand } from "../src/cli/repl.mjs";
import { createTask, listTasks } from "../src/cli/task-store.mjs";
import { createExcelLiveReplFacade } from "../src/excel-live/repl.mjs";
import { initializeExcelConfiguration, resolveExcelStatePaths } from "../src/excel-live/state.mjs";

const workspace = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-test-"));
const sourcePath = path.join(workspace, "source.pdf");
await writeFile(sourcePath, "%PDF-1.7\nOfficeKit task fixture\n");

const session = await createReplSession({ workspaceRoot: workspace, newTaskGoal: "Edit one PDF" });
assert.equal(session.ready.protocol, REPL_PROTOCOL_VERSION);
assert.equal(session.ready.type, "session.ready");
assert.equal(session.ready.task.goal, "Edit one PDF");
assert.equal(session.ready.task.state, "new");
assert.equal(session.ready.resumedFrom, null);
const taskId = session.ready.task.id;

const first = await session.handleLine(JSON.stringify({
  id: "one",
  code: "ctx.state.makeSummary = rows => rows.map(row => row.join(':')).join('|'); console.log('captured', [[1,2]]); return ctx.state.makeSummary([[1,2]]);",
}));
assert.equal(first.protocol, REPL_PROTOCOL_VERSION);
assert.equal(first.ok, true);
assert.equal(first.result, "1:2");
assert.equal(first.events[0].text, "captured [[1,2]]");

const second = await session.handleLine(JSON.stringify({
  id: "two",
  code: "return ctx.state.makeSummary([[3,4]]);",
}));
assert.equal(second.result, "3:4", "ctx.state keeps helpers alive inside one process");

await writeFile(path.join(workspace, "helper.mjs"), "import { FileBlob } from 'office-kit'; export const helper = () => new FileBlob(new Uint8Array([1,2,3]));\n");
const localImport = await session.handleLine(JSON.stringify({
  id: "local-import",
  code: "const helper = await ctx.import('./helper.mjs'); return (await helper.helper()).bytes.length;",
}));
assert.equal(localImport.result, 3);

const failed = await session.handleLine(JSON.stringify({
  id: "failed",
  code: "ctx.state.processOnly = 'not-restored'; throw new Error('controlled failure');",
}));
assert.equal(failed.ok, false);
assert.equal(failed.error.maybeApplied, true);

const staged = await session.handleLine(JSON.stringify({
  id: "input",
  code: `return await ctx.input(${JSON.stringify(sourcePath)}, {artifactId:'main-pdf'});`,
}));
assert.equal(staged.ok, true);
assert.equal(staged.result.artifactId, "main-pdf");
assert.notEqual(staged.result.path, sourcePath);
assert.equal(await readFile(staged.result.path, "utf8"), await readFile(sourcePath, "utf8"));

const commitCell = [
  "const fs = await ctx.import('node:fs/promises');",
  `const bytes = await fs.readFile(${JSON.stringify(staged.result.path)});`,
  "const crypto = await ctx.import('node:crypto');",
  "const sha = crypto.createHash('sha256').update(bytes).digest('hex');",
  "const review = {schemaVersion:1,artifactKind:'pdf',format:'pdf',verdict:'passed-with-limitations',semantic:{status:'passed'},structural:{status:'passed'},layout:{status:'passed'},contentView:{requested:false,status:'not-requested'},visualReview:'unavailable',delivery:{status:'ready',sha256:sha}};",
  "return await ctx.commit(bytes,{artifactId:'main-pdf',summary:'Staged the source as a stable revision',review,next:'Apply the requested edit'});",
].join(" ");
const committed = await session.handleLine(JSON.stringify({ id: "commit", code: commitCell }));
assert.equal(committed.ok, true);
assert.equal(committed.result.commitId, "c0001");
assert.equal(committed.result.type, "officekit.task-commit");

const staleReview = await session.handleLine(JSON.stringify({
  id: "stale-review",
  code: "return await ctx.commit(new TextEncoder().encode('different'),{artifactId:'main-pdf',summary:'bad',review:{schemaVersion:1,artifactKind:'pdf',format:'pdf',verdict:'passed',delivery:{sha256:'0000000000000000000000000000000000000000000000000000000000000000'},visualReview:'complete'}});",
}));
assert.equal(staleReview.ok, false);
assert.equal(staleReview.error.code, "stale-review");
assert.equal(staleReview.error.maybeApplied, false);
assert.equal(session.ctx.task.head.id, "c0001");
assert.ok(session.ctx.task.pending.some((entry) => entry.type === "stale-review"));

const failedReviewCode = [
  "const bytes=new TextEncoder().encode('failed candidate');",
  "const crypto=await ctx.import('node:crypto');",
  "const sha=crypto.createHash('sha256').update(bytes).digest('hex');",
  "return await ctx.commit(bytes,{artifactId:'main-pdf',summary:'Overflowing candidate',review:{schemaVersion:1,artifactKind:'pdf',format:'pdf',verdict:'failed',delivery:{sha256:sha},visualReview:'complete'}});",
].join(" ");
const failedReview = await session.handleLine(JSON.stringify({ id: "failed-review", code: failedReviewCode }));
assert.equal(failedReview.ok, false);
assert.equal(failedReview.error.code, "review-failed");
assert.equal(failedReview.error.maybeApplied, false);
await session.close();

const resumed = await createReplSession({ workspaceRoot: workspace, taskId });
assert.equal(resumed.ready.resumedFrom.commitId, "c0001");
assert.equal(resumed.ready.commit.commitId, "c0001");
assert.equal(resumed.ctx.task.commit.commitId, "c0001");
assert.equal(resumed.ready.artifacts.length, 1);
assert.equal(resumed.ready.artifacts[0].artifactId, "main-pdf");
assert.ok(resumed.ready.task.pending.some((entry) => entry.type === "stale-review"));
assert.ok(resumed.ready.task.pending.some((entry) => entry.type === "review-failed"));
assert.equal(resumed.ready.session.parentSessionId, session.ready.session.id);
const processState = await resumed.handleLine(JSON.stringify({ id: "state", code: "return ctx.state.processOnly ?? null;" }));
assert.equal(processState.result, null, "process-local state is not presented as durable task state");

const published = await resumed.handleLine(JSON.stringify({
  id: "publish",
  code: "return await ctx.publish(ctx.task.commit, {name:'artifact.pdf'});",
}));
assert.equal(published.ok, true);
assert.equal(published.result.reviewVerdict, "passed-with-limitations");
assert.equal(published.result.visualReview, "unavailable");
assert.equal(await readFile(published.result.path, "utf8"), await readFile(sourcePath, "utf8"));
assert.equal((await stat(published.result.path)).isFile(), true);

const rawPublish = await resumed.handleLine(JSON.stringify({
  id: "raw-publish",
  code: "return await ctx.publish(new Uint8Array([1,2,3]),{name:'raw.pdf'});",
}));
assert.equal(rawPublish.ok, false);
assert.equal(rawPublish.error.code, "unreviewed-artifact");
await resumed.close();

const lazyWorkspace = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-lazy-"));
const lazySession = await createReplSession({ workspaceRoot: lazyWorkspace, newTaskGoal: "Lazy state probe" });
const lazyResult = await lazySession.handleLine(JSON.stringify({ id: "lazy", code: "return Object.keys(ctx).sort();" }));
assert.equal(lazyResult.ok, true);
for (const key of ["commit", "input", "publish", "task"]) assert.ok(lazyResult.result.includes(key));
assert.equal(await stat(path.join(lazyWorkspace, ".office-kit")).then(() => true, () => false), true);
assert.equal(await stat(path.join(lazyWorkspace, ".office-kit", "excel")).then(() => true, () => false), false);
assert.equal(await stat(path.join(lazyWorkspace, ".office-kit", "powerpoint")).then(() => true, () => false), false);
await lazySession.close();

const cliWorkspace = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-cli-"));
const input = new PassThrough();
const output = new PassThrough();
const chunks = [];
output.on("data", (chunk) => chunks.push(chunk));
input.end(`${JSON.stringify({ id: "cli", code: "return 7;" })}\n`);
await runReplCommand(["--new", "CLI task", "--workspace", cliWorkspace], { input, output, errorOutput: null });
const lines = Buffer.concat(chunks).toString("utf8").trim().split("\n").map(JSON.parse);
assert.equal(lines[0].type, "session.ready");
assert.equal(lines[1].result, 7);

const failedReadyWorkspace = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-ready-failure-"));
await assert.rejects(
  runReplCommand(["--new", "Recover after output failure", "--workspace", failedReadyWorkspace], {
    input: new PassThrough(),
    output: { write() { throw new Error("output unavailable"); } },
    errorOutput: null,
  }),
  /output unavailable/,
);
const failedReadyTasks = await listTasks({ workspaceRoot: failedReadyWorkspace });
assert.equal(failedReadyTasks.total, 1);
const recoveredAfterOutputFailure = await createReplSession({
  workspaceRoot: failedReadyWorkspace,
  taskId: failedReadyTasks.tasks[0].id,
});
await recoveredAfterOutputFailure.close();

const unsafeOutputWorkspace = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-unsafe-output-"));
const unsafeOutputTask = await createTask({ workspaceRoot: unsafeOutputWorkspace, goal: "Reject an unsafe output root" });
await writeFile(path.join(unsafeOutputWorkspace, "outputs"), "not a directory");
await assert.rejects(
  createReplSession({ workspaceRoot: unsafeOutputWorkspace, taskId: unsafeOutputTask.manifest.id }),
  (error) => error.code === "unsafe-path",
);
await rm(path.join(unsafeOutputWorkspace, "outputs"));
const recoveredAfterInitFailure = await createReplSession({
  workspaceRoot: unsafeOutputWorkspace,
  taskId: unsafeOutputTask.manifest.id,
});
await recoveredAfterInitFailure.close();

const excelRoot = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-excel-"));
const excelPaths = resolveExcelStatePaths({ env: { OFFICEKIT_EXCEL_HOME: excelRoot }, home: excelRoot });
const excelConfig = await initializeExcelConfiguration(excelPaths, { port: 47213 });
const requests = [];
const excel = createExcelLiveReplFacade({
  statePaths: excelPaths,
  platform: "darwin",
  ensureBridge: async () => {},
  doctor: async () => ({ ok: true, result: { host: { status: "ready" } } }),
  bridgeRequestFn: async (_state, method, pathname, body) => {
    requests.push({ method, pathname, body });
    if (pathname === "/v1/cli/sessions") return { ok: true, result: { sessions: [] } };
    if (pathname === "/v1/cli/execute") return { ok: true, result: { value: 2 }, audit: { operation: body.request.operation } };
    return { ok: true, result: { disconnected: body.sessionId } };
  },
});
assert.deepEqual((await excel.sessions()).result.sessions, []);
const request = { protocol: 1, sessionId: "session-1", idempotencyKey: "key-00001", operation: "read_sheets_metadata", args: {} };
assert.equal((await excel.execute(request)).result.value, 2);
assert.equal((await excel.disconnect("session-1")).result.disconnected, "session-1");
assert.equal((await excel.doctor()).result.host.status, "ready");
assert.equal(requests.length, 3);
assert.equal(excelConfig.config.schemaVersion, 1);

console.log("OfficeKit durable REPL smoke ok");
