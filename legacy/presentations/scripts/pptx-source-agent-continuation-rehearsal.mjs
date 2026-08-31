#!/usr/bin/env node

import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { createReplSession } from "../src/cli/repl.mjs";
import { SOURCES } from "./pptx-source-reuse-benchmark.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EVIDENCE_SCHEMA = "office-kit/pptx-source-agent-continuation-rehearsal/v2";
const OVERLAY_IMAGE = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

// This is a deterministic rehearsal of the public Agent path, not a model
// score. It deliberately uses only ctx.input/ctx.commit/ctx.publish and the
// public office-kit package from fresh REPL sessions. Model black-box results
// are recorded separately and preserved when this deterministic evidence is
// regenerated.
export async function runSourceAgentContinuationRehearsal(assetsDir) {
  const results = [];
  for (const source of SOURCES) {
    results.push(await runSourceCase(assetsDir, source));
  }
  return {
    schema: EVIDENCE_SCHEMA,
    protocol: { repl: 2, visualReview: "unavailable", package: "public-office-kit", workflow: "bounded-overlay-resume" },
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
    const firstEdit = await runCell(first, "inspect-reuse-review-commit", firstEditCell({ stagedPath, sourceSlide: source.slide }));
    trace.push(traceEntry(1, firstEdit));
    if (firstEdit.result.failedReview) throw new Error(`${source.id} first review failed: ${JSON.stringify(firstEdit.result.reviewFailure)}`);
    await first.close();

    const resumed = await createReplSession({ workspaceRoot: workspace, taskId });
    const firstRevisionPath = resumed.ready.artifacts[0].path;
    const firstCommit = resumed.ready.commit;
    const secondEdit = await runCell(resumed, "resume-reinspect-review-commit", secondEditCell({ revisionPath: firstRevisionPath, sourceSlideCount: firstEdit.result.sourceSlideCount }));
    trace.push(traceEntry(2, secondEdit));
    if (secondEdit.result.failedReview) throw new Error(`${source.id} resumed review failed: ${JSON.stringify(secondEdit.result.reviewFailure)}`);
    await resumed.close();

    const publishedSession = await createReplSession({ workspaceRoot: workspace, taskId });
    const finalRevisionPath = publishedSession.ready.artifacts[0].path;
    const verification = await runCell(publishedSession, "resume-verify", verifyCell({ revisionPath: finalRevisionPath, sourceSlideCount: firstEdit.result.sourceSlideCount }));
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
      continuationKind: "bounded-overlay",
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

function firstEditCell({ stagedPath, sourceSlide }) {
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
    "const capability=targetSlide.continuationCapability;",
    "if(!capability?.ready||capability.profile!=='bounded-overlay'||capability.embeddedImage!==true) throw new Error('Source-derived slide is not ready for a bounded overlay.');",
    "const textShape=targetSlide.shapes.add({name:'officekit-repl-continuation-text',geometry:'textbox',position:{left:970,top:24,width:140,height:36},fill:'#0F172A',line:{fill:'#0F172A',width:0},text:'OfficeKit',textStyle:{fontFamily:'Arial',fontSize:12,bold:true,color:'#FFFFFF'},accessibility:{title:'OfficeKit continuation marker'}});",
    "const accent=targetSlide.shapes.add({name:'officekit-repl-continuation-accent',geometry:'ellipse',position:{left:1118,top:27,width:30,height:30},fill:'#F97316',line:{fill:'#C2410C',width:1},accessibility:{decorative:true}});",
    `const image=targetSlide.images.add({name:'officekit-repl-continuation-image',alt:'Source-derived continuation image',dataUrl:${JSON.stringify(OVERLAY_IMAGE)},fit:'stretch',position:{left:1154,top:27,width:30,height:30}});`,
    "const target={kind:'bounded-overlay',capability,text:{id:textShape.id,name:textShape.name,value:textShape.text.value},accent:{id:accent.id,name:accent.name,position:accent.position},image:{id:image.id,name:image.name,alt:image.alt}};",
    "const output=await PresentationFile.exportPptx(reopened);",
    "const review=await reviewArtifact(output,{baseline:new FileBlob(cloned.bytes,{type:" + JSON.stringify(PPTX_MIME) + "}),outputPath:path.join(ctx.taskRoot,'candidates','source-continuation-first.pptx'),layout:false,visualReview:'unavailable'}); if(review.verdict==='failed') { const reviewFailure=['semantic','structural','layout','delivery'].flatMap(section=>(review[section]?.issues||[]).filter(issue=>String(issue?.severity||'error').toLowerCase()==='error').map(issue=>({section,kind:issue.kind,type:issue.type,slide:issue.slide,id:issue.id,ids:issue.ids,name:issue.name,names:issue.names,message:issue.message}))).slice(0,32); return {sourceSlideCount,target,review,reviewFailure,failedReview:true}; }",
    "const commit=await ctx.commit(output,{artifactId:'continued-deck',kind:'presentation',name:'continued.pptx',summary:'Reuse a source-derived slide and continue it',review,next:'Reopen the reviewed revision, continue the same page, and verify before publishing'});",
    "return {sourceSlideCount,target,commit,reviewVerdict:review.verdict};",
  ].join(" ");
}

function secondEditCell({ revisionPath, sourceSlideCount }) {
  return [
    "const fs=await ctx.import('node:fs/promises');",
    "const path=await ctx.import('node:path');",
    "const {FileBlob,PresentationFile,reviewArtifact}=await ctx.import('office-kit');",
    `const bytes=await fs.readFile(${JSON.stringify(revisionPath)});`,
    `const baseline=new FileBlob(bytes,{type:${JSON.stringify(PPTX_MIME)}});`,
    `const presentation=await PresentationFile.importPptx(baseline);`,
    `const targetSlide=presentation.slides.items[${sourceSlideCount}];`,
    "const capability=targetSlide.continuationCapability;",
    "if(!capability?.ready||capability.profile!=='bounded-overlay') throw new Error('Reviewed source-derived slide did not retain its continuation capability.');",
    "const textShape=targetSlide.shapes.items.find(candidate=>candidate.name==='officekit-repl-continuation-text');",
    "const accent=targetSlide.shapes.items.find(candidate=>candidate.name==='officekit-repl-continuation-accent');",
    "const image=targetSlide.images.items.find(candidate=>candidate.name==='officekit-repl-continuation-image');",
    "if(!textShape||!accent||!image) throw new Error('Reviewed continuation overlay could not be re-inspected.');",
    "const before={text:textShape.text.value,accent:{...accent.position}};",
    "textShape.text.set('OfficeKit resumed');",
    "accent.position={...accent.position,left:accent.position.left-8,top:accent.position.top+2};",
    "const target={kind:'bounded-overlay',capability,before,text:{id:textShape.id,name:textShape.name,value:textShape.text.value},accent:{id:accent.id,name:accent.name,position:accent.position},image:{id:image.id,name:image.name,alt:image.alt}};",
    "const output=await PresentationFile.exportPptx(presentation);",
    "const review=await reviewArtifact(output,{baseline,outputPath:path.join(ctx.taskRoot,'candidates','source-continuation-resumed.pptx'),layout:false,visualReview:'unavailable'}); if(review.verdict==='failed') { const reviewFailure=['semantic','structural','layout','delivery'].flatMap(section=>(review[section]?.issues||[]).filter(issue=>String(issue?.severity||'error').toLowerCase()==='error').map(issue=>({section,kind:issue.kind,type:issue.type,slide:issue.slide,id:issue.id,ids:issue.ids,name:issue.name,names:issue.names,message:issue.message}))).slice(0,32); return {kind:'bounded-overlay',target,review,reviewFailure,failedReview:true}; }",
    "const commit=await ctx.commit(output,{artifactId:'continued-deck',kind:'presentation',name:'continued.pptx',summary:'Continue the source-derived page after resume',review,next:'Resume once more, verify both edits, and publish'});",
    "return {kind:'bounded-overlay',target,commit,reviewVerdict:review.verdict};",
  ].join(" ");
}

function verifyCell({ revisionPath, sourceSlideCount }) {
  return [
    "const fs=await ctx.import('node:fs/promises');",
    "const {FileBlob,PresentationFile}=await ctx.import('office-kit');",
    `const bytes=await fs.readFile(${JSON.stringify(revisionPath)});`,
    `const presentation=await PresentationFile.importPptx(new FileBlob(bytes,{type:${JSON.stringify(PPTX_MIME)}}));`,
    `const targetSlide=presentation.slides.items[${sourceSlideCount}];`,
    "const capability=targetSlide.continuationCapability;",
    "const textShape=targetSlide.shapes.items.find(candidate=>candidate.name==='officekit-repl-continuation-text');",
    "const accent=targetSlide.shapes.items.find(candidate=>candidate.name==='officekit-repl-continuation-accent');",
    "const image=targetSlide.images.items.find(candidate=>candidate.name==='officekit-repl-continuation-image');",
    "const result={foundResumed:textShape?.text?.value==='OfficeKit resumed',accentMoved:accent?.position?.left===1110&&accent?.position?.top===29,imagePresent:image?.alt==='Source-derived continuation image',capabilityReady:capability?.ready===true,capabilityProfile:capability?.profile};",
    `return {slideCount:presentation.slides.count,sourceSlideCount:${sourceSlideCount},kind:'bounded-overlay',result};`,
  ].join(" ");
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
  process.stdout.write(`${JSON.stringify({ ok: true, output, sources: evidence.sources.length })}\n`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch((error) => {
    process.stderr.write(`${error?.stack || error}\n`);
    process.exitCode = 2;
  });
}
