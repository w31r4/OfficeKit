import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, readdir, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  MAX_AUTHORING_PLAN_BYTES,
  normalizePresentationAuthoringPlan,
} from "../src/cli/authoring-plan.mjs";
import {
  createTask,
  listTasks,
  openTask,
  readTaskPlan,
  stageTaskInput,
  taskDetail,
  writeTaskPlan,
} from "../src/cli/task-store.mjs";
import { formatTaskDetail } from "../src/cli/tasks.mjs";
import { createReplSession } from "../src/cli/repl.mjs";

function validPlan(overrides = {}) {
  return {
    schema: "office-kit/presentation-authoring-plan/v1",
    mode: "create",
    brief: { audience: "Engineering leadership", purpose: "Make one architecture decision" },
    narrative: { thesis: "Adopt the bounded migration path", sections: ["Context", "Decision"] },
    design: {
      sourceMode: "self-directed",
      mechanismPacks: ["technical-architecture"],
      designGrammar: {
        palette: { roles: { background: "#F7F5EF", ink: "#17202A", accent: "#D15B2A" } },
        typography: { title: "Aptos Display", body: "Aptos" },
        densityRhythm: "alternate sparse conclusions with denser evidence",
      },
    },
    pages: [{
      id: "p01-context",
      readerTask: "Understand the current constraint",
      claim: "The current path duplicates deployment work",
      evidence: ["Measured build time"],
      contentBudget: { maxCharacters: 420, maxObjects: 12 },
      compositionIntent: "One dominant comparison with a narrow evidence rail",
    }],
    editorial: { voice: "direct, evidence-led", avoid: ["empty transition phrases"] },
    artifactRefs: [],
    recipe: "tasks/create.md",
    unresolved: [],
    nextAction: "Compose and review the first working draft",
    ...overrides,
  };
}

const first = normalizePresentationAuthoringPlan(validPlan());
const reordered = normalizePresentationAuthoringPlan({
  ...validPlan(),
  editorial: { avoid: ["empty transition phrases"], voice: "direct, evidence-led" },
});
assert.equal(first.sha256, reordered.sha256, "object-key order must not change plan identity");
assert.ok(first.bytes.byteLength < MAX_AUTHORING_PLAN_BYTES);

const cyclic = validPlan();
cyclic.brief.self = cyclic;
assert.throws(() => normalizePresentationAuthoringPlan(cyclic), (error) => error.code === "invalid-authoring-plan");
assert.throws(
  () => normalizePresentationAuthoringPlan(validPlan({ brief: { transform() {} } })),
  (error) => error.code === "invalid-authoring-plan",
);
assert.throws(
  () => normalizePresentationAuthoringPlan(validPlan({ pages: Array.from({ length: 65 }, (_, index) => ({
    id: `p${index + 1}`,
    readerTask: "Read",
    claim: "Claim",
    compositionIntent: "Compose",
  })) })),
  (error) => error.code === "invalid-authoring-plan",
);
assert.throws(
  () => normalizePresentationAuthoringPlan(validPlan({ design: {
    sourceMode: "self-directed",
    mechanismPacks: ["technical-architecture"],
    designGrammar: { rawXml: "<p:sp/>" },
  } })),
  (error) => error.code === "unsafe-authoring-plan",
);
assert.throws(
  () => normalizePresentationAuthoringPlan(validPlan({ design: {
    sourceMode: "self-directed",
    mechanismPacks: ["technical-architecture"],
    designGrammar: { notes: "x".repeat(MAX_AUTHORING_PLAN_BYTES) },
  } })),
  (error) => error.code === "authoring-plan-too-large",
);

const workspace = await mkdtemp(path.join(os.tmpdir(), "officekit-authoring-plan-"));
const created = await createTask({ workspaceRoot: workspace, goal: "Create a planned presentation" });
const source = path.join(workspace, "reference.pptx");
await writeFile(source, "reference-package-bytes");
const staged = await stageTaskInput(created, source, { artifactId: "reference-deck", kind: "presentation" });
const plan = validPlan({
  mode: "create-from-template",
  design: {
    sourceMode: "template",
    mechanismPacks: [],
    designGrammar: { evidence: "Distilled from the staged reference" },
    artifactRef: { artifactId: staged.artifactId, sha256: staged.sha256 },
  },
  artifactRefs: [{ artifactId: staged.artifactId, sha256: staged.sha256, role: "authoritative-template" }],
  recipe: "tasks/create-from-template.md",
});
const descriptor = await writeTaskPlan(created, plan);
assert.equal(descriptor.state, "working");
assert.equal(descriptor.pageCount, 1);
assert.equal(descriptor.unchanged, false);
assert.deepEqual(await readTaskPlan(created), plan);
const planEntries = await readdir(path.join(created.taskRoot, "plans"));
assert.deepEqual(planEntries, [`${descriptor.sha256}.json`]);
const manifestBeforeIdempotentWrite = await readFile(path.join(created.taskRoot, "task.json"));
await assert.rejects(
  writeTaskPlan(created, plan),
  (error) => error.code === "stale-authoring-plan" && error.expectedSha256 === descriptor.sha256,
);
const repeated = await writeTaskPlan(created, plan, { expectedSha256: descriptor.sha256 });
assert.equal(repeated.unchanged, true);
assert.deepEqual(await readFile(path.join(created.taskRoot, "task.json")), manifestBeforeIdempotentWrite);

const changed = structuredClone(plan);
changed.pages[0].claim = "A template-bound plan survives a fresh process";
await assert.rejects(
  writeTaskPlan(created, changed, { expectedSha256: "0".repeat(64) }),
  (error) => error.code === "stale-authoring-plan",
);
const changedDescriptor = await writeTaskPlan(created, changed, { expectedSha256: descriptor.sha256 });
assert.notEqual(changedDescriptor.sha256, descriptor.sha256);
assert.equal(changedDescriptor.state, "working");
assert.equal((await listTasks({ workspaceRoot: workspace })).tasks[0].plan.sha256, changedDescriptor.sha256);
const detail = await taskDetail({ workspaceRoot: workspace, taskId: created.manifest.id });
assert.equal(detail.task.plan.mode, "create-from-template");
assert.doesNotMatch(JSON.stringify(detail.task.plan), /Engineering leadership/u, "task detail must not inline the full plan");
assert.match(formatTaskDetail(detail), /Plan\n  create-from-template · 1 page · tasks\/create-from-template[.]md/u);

const unbound = validPlan({
  artifactRefs: [{ artifactId: "missing", sha256: "1".repeat(64) }],
});
await assert.rejects(
  writeTaskPlan(created, unbound, { expectedSha256: changedDescriptor.sha256 }),
  (error) => error.code === "unbound-authoring-plan-reference",
);

const legacy = await createTask({ workspaceRoot: workspace, goal: "Legacy task" });
const legacyManifestPath = path.join(legacy.taskRoot, "task.json");
const legacyManifest = JSON.parse(await readFile(legacyManifestPath, "utf8"));
legacyManifest.schemaVersion = 1;
delete legacyManifest.plan;
await chmod(legacyManifestPath, 0o600);
await writeFile(legacyManifestPath, `${JSON.stringify(legacyManifest, null, 2)}\n`);
const legacyBytes = await readFile(legacyManifestPath);
assert.equal((await openTask({ workspaceRoot: workspace, taskId: legacy.manifest.id })).manifest.plan, null);
await listTasks({ workspaceRoot: workspace, all: true });
await taskDetail({ workspaceRoot: workspace, taskId: legacy.manifest.id });
assert.deepEqual(await readFile(legacyManifestPath), legacyBytes, "read-only task commands must not migrate schema 1");
const legacyOpened = await openTask({ workspaceRoot: workspace, taskId: legacy.manifest.id });
await stageTaskInput(legacyOpened, source, { artifactId: "legacy-source", kind: "presentation" });
assert.equal(JSON.parse(await readFile(legacyManifestPath, "utf8")).schemaVersion, 2, "the first successful mutation migrates the task");

const replWorkspace = await mkdtemp(path.join(os.tmpdir(), "officekit-authoring-plan-repl-"));
const firstSession = await createReplSession({ workspaceRoot: replWorkspace, newTaskGoal: "Resume a planned deck" });
assert.equal(firstSession.ready.protocol, 3);
assert.equal(firstSession.ctx.protocol, 3);
assert.equal(firstSession.ctx.task.plan, null);
assert.equal(await firstSession.ctx.plan(), null);
const replPlan = validPlan();
const replDescriptor = await firstSession.ctx.plan(replPlan);
assert.equal(replDescriptor.state, "working");
assert.equal(firstSession.ctx.task.plan.sha256, replDescriptor.sha256);
const replTaskId = firstSession.ctx.task.id;
await firstSession.close();

const resumedSession = await createReplSession({ workspaceRoot: replWorkspace, taskId: replTaskId });
assert.equal(resumedSession.ready.protocol, 3);
assert.equal(resumedSession.ready.task.plan.sha256, replDescriptor.sha256);
assert.equal(resumedSession.ready.task.plan.state, "working");
assert.equal("brief" in resumedSession.ready.task.plan, false, "ready envelopes must keep full plan content out of band");
assert.deepEqual(await resumedSession.ctx.plan(), replPlan);
await resumedSession.close();

console.log("presentation authoring plan smoke ok");
