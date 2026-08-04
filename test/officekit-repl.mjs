import assert from "node:assert/strict";
import { appendFile, mkdtemp, readFile, stat, symlink, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  createReplSession,
  REPL_PROTOCOL_VERSION,
  runReplCommand,
} from "../src/cli/repl.mjs";
import { createExcelLiveReplFacade } from "../src/excel-live/repl.mjs";
import { initializeExcelConfiguration, resolveExcelStatePaths } from "../src/excel-live/state.mjs";

const repositoryRoot = path.resolve(import.meta.dirname, "..");
const workspace = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-test-"));
const taskRoot = path.join(workspace, "task");
const session = await createReplSession({ workspaceRoot: workspace, taskRoot });

const first = await session.handleLine(JSON.stringify({
  id: "one",
  code: "ctx.state.makeSummary = rows => rows.map(row => row.join(':')).join('|'); console.log('captured', rows = [[1, 2]]); return ctx.state.makeSummary(rows);",
}));
assert.equal(first.protocol, REPL_PROTOCOL_VERSION);
assert.equal(first.ok, true);
assert.equal(first.result, "1:2");
assert.equal(first.events[0].text, "captured [[1,2]]");

const second = await session.handleLine(JSON.stringify({
  id: "two",
  code: "return ctx.state.makeSummary([[3, 4]]);",
}));
assert.equal(second.result, "3:4", "ctx.state keeps helpers alive between cells");

await writeFile(path.join(workspace, "helper.mjs"), "import { FileBlob } from 'office-kit'; export const helper = () => new FileBlob(new Uint8Array([1, 2, 3]));\n");
const localImport = await session.handleLine(JSON.stringify({
  id: "local-import",
  code: "const helper = await ctx.import('./helper.mjs'); return (await helper.helper()).bytes.length;",
}));
assert.equal(localImport.result, 3, "workspace helpers can import the published OfficeKit package");

const failed = await session.handleLine(JSON.stringify({
  id: "failed",
  code: "ctx.state.afterFailure = 'kept'; throw new Error('controlled failure');",
}));
assert.equal(failed.ok, false);
assert.equal(failed.error.code, "execution-failed");
assert.equal(failed.error.maybeApplied, true);
const afterFailure = await session.handleLine(JSON.stringify({
  id: "after-failure",
  code: "return ctx.state.afterFailure;",
}));
assert.equal(afterFailure.result, "kept");

const badJson = await session.handleLine("not-json");
assert.equal(badJson.ok, false);
assert.equal(badJson.error.code, "invalid-json");
const privateImport = await session.handleLine(JSON.stringify({
  id: "private-import",
  code: "return await ctx.import('office-kit/src/index.mjs');",
}));
assert.equal(privateImport.ok, false);
assert.equal(privateImport.error.code, "unpublished-subpath");
assert.equal(privateImport.error.maybeApplied, false);
const remoteImport = await session.handleLine(JSON.stringify({
  id: "remote-import",
  code: "return await ctx.import('https://example.test/task.mjs');",
}));
assert.equal(remoteImport.error.code, "remote-import");

const published = await session.handleLine(JSON.stringify({
  id: "publish",
  code: "const {FileBlob} = await ctx.import('office-kit'); return await ctx.publish(new FileBlob(new TextEncoder().encode('artifact'), {type: 'application/pdf'}), {name: 'artifact.pdf', kind: 'pdf'});",
}));
assert.equal(published.ok, true);
assert.equal(published.result.kind, "pdf");
assert.equal(published.result.bytes, 8);
assert.match(published.result.sha256, /^[a-f0-9]{64}$/u);
assert.equal(await readFile(published.result.path, "utf8"), "artifact");
const outputStat = await stat(published.result.path);
assert.equal(outputStat.isFile(), true);

const protectedInput = path.join(workspace, "input.bin");
await (await import("node:fs/promises")).writeFile(protectedInput, "keep");
const overwrite = await session.handleLine(JSON.stringify({
  id: "overwrite",
  code: `return await ctx.publish(${JSON.stringify(protectedInput)}, {path: ${JSON.stringify(protectedInput)}, sourcePaths: [${JSON.stringify(protectedInput)}]});`,
}));
assert.equal(overwrite.ok, false);
assert.equal(overwrite.error.code, "unsafe-output");
assert.equal(await readFile(protectedInput, "utf8"), "keep");

const evidenceFile = path.join(taskRoot, "evidence", "qa.txt");
await (await import("node:fs/promises")).writeFile(evidenceFile, "qa");
const evidence = await session.handleLine(JSON.stringify({
  id: "evidence",
  code: "return await ctx.recordEvidence('qa.txt', {kind: 'render', locator: {page: 1}, visualReview: 'unavailable'});",
}));
assert.equal(evidence.ok, true);
assert.equal(evidence.result.visualReview, "unavailable");
assert.equal(evidence.result.bytes, 2);
await session.close();

const checkpoint = path.join(taskRoot, ".officekit-repl");
const checkpointDirectory = (await (await import("node:fs/promises")).readdir(checkpoint)).find(Boolean);
assert.ok(checkpointDirectory);
const checkpointPath = path.join(checkpoint, checkpointDirectory, "checkpoint.json");
const checkpointValue = JSON.parse(await readFile(checkpointPath, "utf8"));
assert.equal(checkpointValue.last.source, "return await ctx.recordEvidence('qa.txt', {kind: 'render', locator: {page: 1}, visualReview: 'unavailable'});");
assert.equal(checkpointValue.state.safe.afterFailure, "kept");

await appendFile(path.join(path.dirname(checkpointPath), "session.jsonl"), `${JSON.stringify({
  protocol: 1,
  type: "request.started",
  sessionId: checkpointValue.sessionId,
  sequence: checkpointValue.sequence + 1,
  id: "interrupted",
  source: "return 99",
  sourceSha256: "interrupted-source",
})}\n`);

const resumed = await createReplSession({ resume: checkpointPath });
const resumedResult = await resumed.handleLine(JSON.stringify({
  id: "resume",
  code: "return {state: ctx.state.afterFailure, artifacts: ctx.publish ? 'ready' : 'missing'};",
}));
assert.equal(resumedResult.ok, true);
assert.equal(resumedResult.result.state, "kept");
assert.equal(resumedResult.result.artifacts, "ready");
assert.equal(resumedResult.audit.maybeApplied, true);
assert.equal(resumedResult.audit.interruptedRequest.id, "interrupted");
await resumed.close();

const unsafeWorkspace = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-symlink-"));
const escapedOutputs = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-escaped-"));
await symlink(escapedOutputs, path.join(unsafeWorkspace, "outputs"));
await assert.rejects(
  () => createReplSession({ workspaceRoot: unsafeWorkspace, taskRoot: path.join(unsafeWorkspace, "task") }),
  (error) => error.code === "unsafe-path",
);

const lazyWorkspace = await mkdtemp(path.join(os.tmpdir(), "officekit-repl-lazy-"));
const lazySession = await createReplSession({ workspaceRoot: lazyWorkspace, taskRoot: path.join(lazyWorkspace, "task") });
const lazyResult = await lazySession.handleLine(JSON.stringify({ id: "lazy", code: "return Object.keys(ctx).sort();" }));
assert.equal(lazyResult.ok, true);
const lazyExcelState = await (await import("node:fs/promises")).access(path.join(lazyWorkspace, "task", "excel")).then(() => true, () => false);
assert.equal(lazyExcelState, false);
await lazySession.close();

const commandInput = [
  JSON.stringify({ id: "cli", code: "return await new Promise(resolve => setTimeout(() => resolve(7), 1));" }),
].join("\n") + "\n";
const output = new (await import("node:stream")).PassThrough();
const input = new (await import("node:stream")).PassThrough();
const chunks = [];
output.on("data", (chunk) => chunks.push(chunk));
input.end(commandInput);
await runReplCommand(["--workspace", workspace, "--task-root", path.join(workspace, "cli-task")], { input, output, errorOutput: null });
assert.equal(JSON.parse(Buffer.concat(chunks).toString("utf8")).result, 7);

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
const request = {
  protocol: 1,
  sessionId: "session-1",
  idempotencyKey: "key-00001",
  operation: "read_sheets_metadata",
  args: {},
};
assert.equal((await excel.execute(request)).result.value, 2);
assert.equal((await excel.disconnect("session-1")).result.disconnected, "session-1");
assert.equal(requests.length, 3);
assert.equal((await excel.doctor()).result.host.status, "ready");
assert.equal(excelConfig.config.schemaVersion, 1);
const unsupportedExcel = createExcelLiveReplFacade({ statePaths: excelPaths, platform: "linux" });
await assert.rejects(() => unsupportedExcel.sessions(), (error) => error.code === "unsupported-platform");

console.log("OfficeKit REPL smoke ok");
