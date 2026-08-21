#!/usr/bin/env node

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { createReplSession } from "../src/cli/repl.mjs";
import { SOURCES } from "./pptx-source-reuse-benchmark.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-source-agent-continuation-rehearsal/v1";

// This is a deterministic rehearsal of the public Agent path, not a model
// score. It deliberately uses only ctx.input/ctx.commit/ctx.publish and the
// public office-kit package from fresh REPL sessions. The separate model
// black-box 3/3 acceptance remains an external Goal gate.
export async function runSourceAgentContinuationRehearsal(assetsDir) {
  const results = [];
  for (const source of SOURCES) {
    results.push(await runSourceCase(assetsDir, source));
  }
  return {
    schema: EVIDENCE_SCHEMA,
    protocol: { repl: 2, visualReview: "unavailable", package: "public-office-kit" },
    modelBlackBox: { required: 3, completed: 0, status: "open" },
    sources: results,
  };
}

async function runSourceCase(assetsDir, source) {
  const sourcePath = path.join(assetsDir, source.fileName);
  const sourceBytes = await readFile(sourcePath);
  const sourceSha256 = sha256(sourceBytes);
  const workspace = await mkdtemp(path.join(os.tmpdir(), "officekit-pptx-source-agent-"));
  const trace = [];
  try {
    const first = await createReplSession({
      workspaceRoot: workspace,
      newTaskGoal: `Continue the ${source.id} presentation without rebuilding its source graph`,
    });
    const taskId = first.ready.task.id;
    const stage = await runCell(first, "stage-input", `return await ctx.input(${JSON.stringify(sourcePath)}, {artifactId:"source-deck"});`);
    trace.push(traceEntry(1, stage));
    const stagedPath = stage.result.path;
    const firstEdit = await runCell(first, "inspect-reuse-review-commit", firstEditCell({ stagedPath, sourceSlide: source.slide, kind: continuationKind(source.id) }));
    trace.push(traceEntry(1, firstEdit));
    if (firstEdit.result.failedReview) throw new Error(`${source.id} first review failed: ${JSON.stringify(firstEdit.result.review).slice(0, 8000)}`);
    await first.close();

    const resumed = await createReplSession({ workspaceRoot: workspace, taskId });
    const firstRevisionPath = resumed.ready.artifacts[0].path;
    const firstCommit = resumed.ready.commit;
    const secondEdit = await runCell(resumed, "resume-reinspect-review-commit", secondEditCell({ revisionPath: firstRevisionPath, sourceSlideCount: firstEdit.result.sourceSlideCount, kind: continuationKind(source.id) }));
    trace.push(traceEntry(2, secondEdit));
    if (secondEdit.result.failedReview) throw new Error(`${source.id} resumed review failed: ${JSON.stringify(secondEdit.result.review).slice(0, 8000)}`);
    await resumed.close();

    const publishedSession = await createReplSession({ workspaceRoot: workspace, taskId });
    const finalRevisionPath = publishedSession.ready.artifacts[0].path;
    const verification = await runCell(publishedSession, "resume-verify", verifyCell({ revisionPath: finalRevisionPath, sourceSlideCount: firstEdit.result.sourceSlideCount, kind: continuationKind(source.id) }));
    trace.push(traceEntry(3, verification));
    const published = await runCell(publishedSession, "publish-reviewed-deck", `return await ctx.publish(ctx.task.commit, {artifactId:"continued-deck", name:${JSON.stringify(`${source.id}-continued.pptx`)}});`);
    trace.push(traceEntry(3, published));
    await publishedSession.close();

    const publishedBytes = await readFile(published.result.path);
    assert.deepEqual(await readFile(sourcePath), sourceBytes, `${source.id} source changed during rehearsal`);
    return {
      id: source.id,
      fileName: source.fileName,
      sourceSha256,
      taskId: "t_<redacted>",
      taskIdValidated: /^t_[0-9a-f]{12}$/u.test(taskId),
      sourceSlide: source.slide,
      sourceSlideCount: firstEdit.result.sourceSlideCount,
      continuationKind: continuationKind(source.id),
      commits: [firstCommit, secondEdit.result.commit].map((commit) => ({
        commitId: commit.commitId,
        revisionSha256: commit.revisionSha256,
        reviewVerdict: commit.reviewVerdict,
        visualReview: commit.visualReview,
      })),
      finalVerification: verification.result,
      publishedSha256: sha256(publishedBytes),
      publishedPathRelative: path.posix.join("outputs", path.basename(published.result.path)),
      sourceUnchanged: true,
      freshSessions: 3,
      trace,
    };
  } finally {
    await rm(workspace, { recursive: true, force: true });
  }
}

function firstEditCell({ stagedPath, sourceSlide, kind }) {
  return [
    "const fs=await ctx.import('node:fs/promises');",
    "const path=await ctx.import('node:path');",
    "const {FileBlob,PresentationFile,reviewArtifact}=await ctx.import('office-kit');",
    `const bytes=await fs.readFile(${JSON.stringify(stagedPath)});`,
    `const presentation=await PresentationFile.importPptx(new FileBlob(bytes,{type:${JSON.stringify(PPTX_MIME)}}));`,
    "const sourceSlideCount=presentation.slides.count;",
    `const sourceSlide=presentation.slides.items[${sourceSlide - 1}];`,
    "const clone=sourceSlide.duplicate();",
    "clone.moveTo(sourceSlideCount);",
    "const cloned=await PresentationFile.exportPptx(presentation);",
    "const reopened=await PresentationFile.importPptx(cloned.bytes);",
    "const targetSlide=reopened.slides.items[sourceSlideCount];",
    `const target=${continuationMutation("targetSlide", "first", kind)};`,
    "const output=await PresentationFile.exportPptx(reopened);",
    "const review=await reviewArtifact(output,{baseline:new FileBlob(cloned.bytes,{type:" + JSON.stringify(PPTX_MIME) + "}),outputPath:path.join(ctx.taskRoot,'candidates','source-continuation-first.pptx'),layout:false,visualReview:'unavailable'}); if(review.verdict==='failed') return {sourceSlideCount,target,review,failedReview:true};",
    "const commit=await ctx.commit(output,{artifactId:'continued-deck',kind:'presentation',name:'continued.pptx',summary:'Reuse a source-derived slide and continue it',review,next:'Reopen the reviewed revision, continue the same page, and verify before publishing'});",
    "return {sourceSlideCount,target,commit,reviewVerdict:review.verdict};",
  ].join(" ");
}

function secondEditCell({ revisionPath, sourceSlideCount, kind }) {
  return [
    "const fs=await ctx.import('node:fs/promises');",
    "const path=await ctx.import('node:path');",
    "const {FileBlob,PresentationFile,reviewArtifact}=await ctx.import('office-kit');",
    `const bytes=await fs.readFile(${JSON.stringify(revisionPath)});`,
    `const baseline=new FileBlob(bytes,{type:${JSON.stringify(PPTX_MIME)}});`,
    `const presentation=await PresentationFile.importPptx(baseline);`,
    `const targetSlide=presentation.slides.items[${sourceSlideCount}];`,
    `const target=${continuationMutation("targetSlide", "resumed", kind)};`,
    "const output=await PresentationFile.exportPptx(presentation);",
    "const review=await reviewArtifact(output,{baseline,outputPath:path.join(ctx.taskRoot,'candidates','source-continuation-resumed.pptx'),layout:false,visualReview:'unavailable'}); if(review.verdict==='failed') return {kind:" + JSON.stringify(kind) + ",target,review,failedReview:true};",
    "const commit=await ctx.commit(output,{artifactId:'continued-deck',kind:'presentation',name:'continued.pptx',summary:'Continue the source-derived page after resume',review,next:'Resume once more, verify both edits, and publish'});",
    `return {kind:${JSON.stringify(kind)},target,commit,reviewVerdict:review.verdict};`,
  ].join(" ");
}

function verifyCell({ revisionPath, sourceSlideCount, kind }) {
  return [
    "const fs=await ctx.import('node:fs/promises');",
    "const {FileBlob,PresentationFile}=await ctx.import('office-kit');",
    `const bytes=await fs.readFile(${JSON.stringify(revisionPath)});`,
    `const presentation=await PresentationFile.importPptx(new FileBlob(bytes,{type:${JSON.stringify(PPTX_MIME)}}));`,
    `const targetSlide=presentation.slides.items[${sourceSlideCount}];`,
    `const result=${verificationExpression("targetSlide", kind)};`,
    `return {slideCount:presentation.slides.count,sourceSlideCount:${sourceSlideCount},kind:${JSON.stringify(kind)},result};`,
  ].join(" ");
}

function continuationMutation(slideExpression, phase, kind) {
  // The rehearsal must exercise a semantic text edit without inventing a
  // longer title that creates a new overflow finding.  Real tasks still use
  // the review gate to reject content that does not fit.
  return `(() => { const slide=${slideExpression}; const kind=${JSON.stringify(kind)}; const textShape=slide.shapes.items.find(candidate=>candidate.text?.value); const image=slide.images.items.find(candidate=>candidate.dataUrl?.startsWith('data:image/svg+xml;base64,')); if(kind==='text' && textShape) { const before=textShape.text.value; const value=${JSON.stringify(phase === "first" ? "OK" : "OK·R")}; textShape.text.set(value); return {kind:'text',before,value}; } if(kind==='svg-text' && image) { const leaf=image.getSvgTextNodes()[0]; if(!leaf) throw new Error('No editable SVG text leaf found.'); const value=leaf.text+${JSON.stringify(phase === "first" ? " · A" : " · B")}; const edit=image.editSvgText(leaf.id,{expectedHash:leaf.expectedHash,value}); return {kind:'svg-text',nodeId:leaf.id,before:leaf.text,value,expectedHash:edit.expectedHash,sourceSha256:edit.sourceSha256}; } throw new Error('No supported continuation leaf found.'); })()`;
}

function verificationExpression(slideExpression, kind) {
  if (kind === "text") {
    return `(() => { const slide=${slideExpression}; const values=slide.shapes.items.map(candidate=>candidate.text?.value||''); return {foundResumed:values.some(value=>value==='OK·R'),textShapes:values.filter(Boolean).length}; })()`;
  }
  return `(() => { const slide=${slideExpression}; const images=slide.images.items.filter(candidate=>candidate.dataUrl?.startsWith('data:image/svg+xml;base64,')); const values=images.flatMap(candidate=>candidate.getSvgTextNodes().map(node=>node.text)); return {foundResumed:values.some(value=>value.endsWith(' · B')),svgImages:images.length}; })()`;
}

function continuationKind(id) {
  return id === "mckinsey-customer-loyalty" ? "svg-text" : "text";
}

async function runCell(session, id, code) {
  const response = await session.handleLine(JSON.stringify({ id, code }));
  if (!response.ok) throw new Error(`${id} failed: ${response.error?.message || "unknown REPL error"}; response=${JSON.stringify(response).slice(0, 4000)}`);
  return response;
}

function traceEntry(session, response) {
  return {
    session,
    id: response.id,
    ok: response.ok,
    maybeApplied: response.audit?.maybeApplied ?? false,
    imports: response.imports || [],
  };
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function parseArgs(argv) {
  let assetsDir;
  let output;
  let force = false;
  for (let index = 0; index < argv.length; index += 1) {
    const flag = argv[index];
    if (flag === "--assets-dir") assetsDir = argv[++index];
    else if (flag === "--output") output = argv[++index];
    else if (flag === "--force") force = true;
    else throw new Error(`Unknown option ${flag}.`);
  }
  if (!assetsDir || !output) throw new Error("Usage: pptx-source-agent-continuation-rehearsal.mjs --assets-dir <dir> --output <evidence.json> [--force]");
  return { assetsDir: path.resolve(assetsDir), output: path.resolve(output), force };
}

async function main() {
  const { assetsDir, output, force } = parseArgs(process.argv.slice(2));
  const evidence = await runSourceAgentContinuationRehearsal(assetsDir);
  await writeFile(output, `${JSON.stringify(evidence, null, 2)}\n`, { flag: force ? "w" : "wx" });
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length, modelBlackBox: evidence.modelBlackBox })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
