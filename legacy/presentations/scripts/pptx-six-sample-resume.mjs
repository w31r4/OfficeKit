#!/usr/bin/env node

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { createReplSession } from "../src/cli/repl.mjs";
import { SOURCES } from "./pptx-six-sample-import.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-six-sample-resume-evidence/v1";
const DEFAULT_ASSETS_DIR = path.resolve("tmp/reference-pptx-downloads");
const DEFAULT_OUTPUT = path.resolve("tmp/presentation-six-sample-import/resume-evidence.v1.json");
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;

// This is an optional evidence rehearsal, not a routine test.  It uses the
// public REPL contract from three fresh sessions and keeps only hashes and
// compact commit data; task directories and reference inputs stay disposable.
export async function collectSixSampleResumeEvidence({ assetsDir = DEFAULT_ASSETS_DIR } = {}) {
  const results = [];
  for (const source of SOURCES) results.push(await runSourceCase(path.resolve(assetsDir), source));
  return {
    schema: EVIDENCE_SCHEMA,
    protocol: { repl: 2, visualReview: "unavailable", package: "public-office-kit" },
    sources: results,
  };
}

async function runSourceCase(assetsDir, source) {
  const sourcePath = path.join(assetsDir, source.fileName);
  const sourceBytes = await readBounded(sourcePath);
  const sourceSha256 = sha256(sourceBytes);
  assert.equal(sourceSha256, source.sha256, `${source.id}: source SHA-256 mismatch`);
  const workspace = await mkdtemp(path.join(os.tmpdir(), "officekit-six-sample-resume-"));
  try {
    const firstSession = await createReplSession({
      workspaceRoot: workspace,
      newTaskGoal: `Continue imported ${source.id} without rebuilding its source graph`,
    });
    const taskId = firstSession.ready.task.id;
    const staged = await firstSession.ctx.input(sourcePath, { artifactId: "source-deck" });
    const first = await runCell(firstSession, "edit-before-resume", firstEditCell(staged.path));
    await firstSession.close();

    const secondSession = await createReplSession({ workspaceRoot: workspace, taskId });
    const firstRevision = secondSession.ready.artifacts.find((artifact) => artifact.artifactId === "continued-deck");
    if (!firstRevision) throw new Error(`${source.id}: reviewed revision was not restored`);
    const second = await runCell(secondSession, "edit-after-resume", secondEditCell(firstRevision.path, first.result));
    await secondSession.close();

    const finalSession = await createReplSession({ workspaceRoot: workspace, taskId });
    const finalRevision = finalSession.ready.artifacts.find((artifact) => artifact.artifactId === "continued-deck");
    if (!finalRevision) throw new Error(`${source.id}: second reviewed revision was not restored`);
    const verification = await runCell(finalSession, "verify-after-second-resume", verifyCell(finalRevision.path, second.result));
    const published = await runCell(finalSession, "publish-reviewed-revision", "return await ctx.publish(ctx.task.commit,{artifactId:'continued-deck',name:'continued.pptx'});");
    const publishedBytes = await readFile(published.result.path);
    const sourceAfter = await readFile(sourcePath);
    await finalSession.close();
    assert.deepEqual(sourceAfter, sourceBytes, `${source.id}: source changed during resume rehearsal`);
    return {
      id: source.id,
      fileName: source.fileName,
      sourceSha256,
      taskId: "t_<redacted>",
      taskIdValidated: /^t_[0-9a-f]{12}$/u.test(taskId),
      sessions: 3,
      commits: [first.result.commit, second.result.commit].map(commitSummary),
      verification: verification.result,
      publishedSha256: sha256(publishedBytes),
      sourceUnchanged: true,
    };
  } finally {
    await rm(workspace, { recursive: true, force: true });
  }
}

function firstEditCell(stagedPath) {
  return [
    "const fs=await ctx.import('node:fs/promises');",
    "const path=await ctx.import('node:path');",
    "const {FileBlob,PresentationFile,reviewArtifact}=await ctx.import('office-kit');",
    `const bytes=await fs.readFile(${JSON.stringify(stagedPath)});`,
    `const baseline=new FileBlob(bytes,{type:${JSON.stringify(PPTX_MIME)}});`,
    "const presentation=await PresentationFile.importPptx(baseline);",
    "const shape=presentation.slides.items.flatMap(slide=>slide.shapes.items).find(candidate=>candidate.text?.paragraphs?.some(paragraph=>paragraph.runs?.some(run=>typeof run.text==='string'&&run.text.trim())));",
    "if(!shape) throw new Error('no top-level text shape for resume evidence');",
    "const run=shape.text.paragraphs.flatMap(paragraph=>paragraph.runs||[]).find(candidate=>typeof candidate.text==='string'&&candidate.text.trim());",
    "const before=run.text; const after=before.length>1?'X'+before.slice(1):'X'; shape.text.replace(before,after);",
    "const output=await PresentationFile.exportPptx(presentation);",
    "const review=await reviewArtifact(output,{baseline,outputPath:path.join(ctx.taskRoot,'candidates','resume-first.pptx'),layout:false,visualReview:'unavailable'});",
    "if(review.verdict==='failed') throw new Error(JSON.stringify(review));",
    "const commit=await ctx.commit(output,{artifactId:'continued-deck',kind:'presentation',name:'continued.pptx',summary:'Continue imported deck with a bounded text edit',review,next:'Resume and re-inspect the reviewed revision'});",
    "return {targetId:shape.id,before,after,commit};",
  ].join(" ");
}

function secondEditCell(revisionPath, first) {
  return [
    "const fs=await ctx.import('node:fs/promises');",
    "const path=await ctx.import('node:path');",
    "const {FileBlob,PresentationFile,reviewArtifact}=await ctx.import('office-kit');",
    `const bytes=await fs.readFile(${JSON.stringify(revisionPath)});`,
    `const baseline=new FileBlob(bytes,{type:${JSON.stringify(PPTX_MIME)}});`,
    "const presentation=await PresentationFile.importPptx(baseline);",
    `const shape=presentation.resolve(${JSON.stringify(first.targetId)});`,
    "if(!shape?.text) throw new Error('source-bound text target was not restored after resume');",
    `const before=${JSON.stringify(first.after)};`,
    "const after=before.length>1?'Y'+before.slice(1):'Y'; shape.text.replace(before,after);",
    "const output=await PresentationFile.exportPptx(presentation);",
    "const review=await reviewArtifact(output,{baseline,outputPath:path.join(ctx.taskRoot,'candidates','resume-second.pptx'),layout:false,visualReview:'unavailable'});",
    "if(review.verdict==='failed') throw new Error(JSON.stringify(review));",
    "const commit=await ctx.commit(output,{artifactId:'continued-deck',kind:'presentation',name:'continued.pptx',summary:'Continue the same imported target after resume',review,next:'Resume once more and verify both edits'});",
    "return {targetId:shape.id,before,after,commit};",
  ].join(" ");
}

function verifyCell(revisionPath, second) {
  return [
    "const fs=await ctx.import('node:fs/promises');",
    "const {createHash}=await ctx.import('node:crypto');",
    "const {FileBlob,PresentationFile}=await ctx.import('office-kit');",
    `const bytes=await fs.readFile(${JSON.stringify(revisionPath)});`,
    `const presentation=await PresentationFile.importPptx(new FileBlob(bytes,{type:${JSON.stringify(PPTX_MIME)}}));`,
    `const shape=presentation.resolve(${JSON.stringify(second.targetId)});`,
    `const expected=${JSON.stringify(second.after)};`,
    "const observed=shape?.text?.value||''; const digest=value=>createHash('sha256').update(value).digest('hex');",
    "return {slideCount:presentation.slides.count,targetId:shape?.id||null,observedLength:observed.length,observedTextSha256:digest(observed),expectedLength:expected.length,expectedTextSha256:digest(expected),found:observed.includes(expected)};",
  ].join(" ");
}

async function runCell(session, id, code) {
  const response = await session.handleLine(JSON.stringify({ id, code }));
  if (!response.ok) throw new Error(`${id} failed: ${response.error?.message || "unknown REPL error"}`);
  return response;
}

function commitSummary(commit) {
  return {
    commitId: commit.commitId,
    revisionSha256: commit.revisionSha256,
    reviewVerdict: commit.reviewVerdict,
    visualReview: commit.visualReview,
  };
}

async function readBounded(filePath) {
  const info = await stat(filePath);
  if (!info.isFile() || info.size < 1 || info.size > MAX_SOURCE_BYTES) throw new RangeError(`PPTX input is outside 1..${MAX_SOURCE_BYTES}: ${filePath}`);
  const bytes = await readFile(filePath);
  if (bytes.byteLength !== info.size) throw new Error(`PPTX input changed while reading: ${filePath}`);
  return bytes;
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function parseArgs(argv) {
  let assetsDir = DEFAULT_ASSETS_DIR;
  let output = DEFAULT_OUTPUT;
  let force = false;
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === "--assets-dir") assetsDir = argv[++index];
    else if (argv[index] === "--output") output = argv[++index];
    else if (argv[index] === "--force") force = true;
    else throw new Error(`Unknown option ${argv[index]}.`);
  }
  return { assetsDir, output, force };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const evidence = await collectSixSampleResumeEvidence(options);
  const output = path.resolve(options.output);
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: options.force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
