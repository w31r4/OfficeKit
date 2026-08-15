import assert from "node:assert/strict";
import { chmod, mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  DocumentFile,
  DocumentModel,
  PdfArtifact,
  PdfFile,
  Presentation,
  PresentationFile,
  reviewArtifact,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";
import { createReplSession } from "../src/cli/repl.mjs";

const workspace = await mkdtemp(path.join(os.tmpdir(), "officekit-task-artifacts-"));
const source = path.join(workspace, "source.docx");
await (await DocumentFile.exportDocx(DocumentModel.create({ paragraphs: ["Original input"] }))).save(source);

const session = await createReplSession({ workspaceRoot: workspace, newTaskGoal: "Create four related deliverables" });
const taskId = session.ready.task.id;
const staged = await session.ctx.input(source, { artifactId: "source-document" });
await writeFile(source, "externally changed bytes");

const document = await DocumentFile.exportDocx(DocumentModel.create({ paragraphs: ["Committed DOCX"] }));
const workbook = Workbook.create();
workbook.worksheets.add("Data").getRange("A1:B2").values = [["Metric", "Value"], ["Revenue", 42]];
const spreadsheet = await SpreadsheetFile.exportXlsx(workbook);
const presentationModel = Presentation.create();
presentationModel.slides.add({ name: "Result" }).shapes.add({
  geometry: "textbox",
  text: "Committed PPTX",
  position: { left: 20, top: 20, width: 400, height: 80 },
});
const presentationSource = await PresentationFile.exportPptx(presentationModel);
const importedPresentation = await PresentationFile.importPptx(presentationSource);
importedPresentation.slides.getItem(0).shapes.getItemAt(0).text.set("Committed PPTX through an Edit Plan");
const presentation = await PresentationFile.exportPptx(importedPresentation);
assert.equal(presentation.metadata.editPlan?.schema, "office-kit/pptx-edit-plan/v1");
const pdf = await PdfFile.exportPdf(PdfArtifact.create({ pages: [{ text: "Committed PDF" }] }));

let head;
let firstCommit;
let presentationCommit;
for (const [artifactId, kind, name, value] of [
  ["final-document", "document", "final.docx", document],
  ["final-workbook", "workbook", "final.xlsx", spreadsheet],
  ["final-presentation", "presentation", "final.pptx", presentation],
  ["final-pdf", "pdf", "final.pdf", pdf],
]) {
  const review = await reviewArtifact(value, {
    outputPath: path.join(workspace, "candidates", name),
    layout: false,
    visualReview: "unavailable",
  });
  assert.notEqual(review.verdict, "failed", JSON.stringify(review, null, 2));
  head = await session.ctx.commit(value, {
    artifactId,
    kind,
    name,
    summary: `Committed ${name}`,
    review,
    next: artifactId === "final-pdf" ? "Publish all four files" : "Continue creating deliverables",
  });
  firstCommit ??= head;
  if (artifactId === "final-presentation") presentationCommit = head;
}
assert.equal(head.commitId, "c0004");
assert.equal(head.artifacts.length, 4);
assert.equal(presentationCommit.operation?.schema, "office-kit/task-edit-plan/v1");
const operationRecord = JSON.parse(await readFile(path.join(session.ctx.taskRoot, presentationCommit.operation.path), "utf8"));
assert.equal(operationRecord.plan.outputSha256, presentationCommit.revisionSha256);
assert.equal(operationRecord.plan.operations.length, 1);
assert.equal(operationRecord.plan.operations[0].leafKind, "text");
assert.equal(operationRecord.plan.operations[0].footprint.leafKind, "text");
await session.close();

const resumed = await createReplSession({ workspaceRoot: workspace, taskId });
assert.equal(resumed.ready.resumedFrom.commitId, "c0004");
assert.equal(resumed.ready.commit.commitId, "c0004");
assert.equal(resumed.ready.artifacts.length, 4);
assert.equal(resumed.ready.operations.length, 1);
assert.equal(resumed.ready.operations[0].commitId, presentationCommit.commitId);
assert.equal(resumed.ready.operations[0].operationIds.length, 1);
assert.equal(path.isAbsolute(resumed.ready.operations[0].path), true);
assert.ok(resumed.ready.task.pending.some((entry) => entry.type === "source-changed" && entry.artifactId === staged.artifactId));
await assert.rejects(
  resumed.ctx.publish(firstCommit, { name: "stale.docx" }),
  (error) => error.code === "stale-commit",
);
for (const artifact of resumed.ready.artifacts) {
  const published = await resumed.ctx.publish(resumed.ready.commit, {
    artifactId: artifact.artifactId,
    name: artifact.name,
  });
  assert.equal(published.sha256, artifact.sha256);
  assert.equal(published.reviewVerdict, "passed-with-limitations");
  assert.deepEqual(await readFile(published.path), await readFile(artifact.path));
}
const storedOperationPath = resumed.ready.operations[0].path;
await resumed.close();
await chmod(storedOperationPath, 0o600);
await writeFile(storedOperationPath, "{}\n");
await assert.rejects(
  createReplSession({ workspaceRoot: workspace, taskId }),
  (error) => error.code === "operation-corrupt",
);

// A fresh Agent context must be able to continue a source-bound PPTX edit from
// reviewed bytes and immutable Edit Plan evidence, without restoring the old
// JavaScript heap. Use the real SmartArt canary so the first commit mutates a
// dependent DiagramDataPart and the second commit mutates a SlidePart.
const strategyWorkspace = await mkdtemp(path.join(os.tmpdir(), "officekit-task-pptx-resume-"));
const strategySourcePath = path.resolve("evals/assets/presentations/strategy-review.pptx");
const strategySourceBytes = await readFile(strategySourcePath);
const strategySession = await createReplSession({
  workspaceRoot: strategyWorkspace,
  newTaskGoal: "Edit and review the strategy presentation without rebuilding it",
});
const strategyTaskId = strategySession.ready.task.id;
const strategyInput = await strategySession.handleLine(JSON.stringify({
  id: "stage-strategy-source",
  code: `return await ctx.input(${JSON.stringify(strategySourcePath)}, {artifactId:'strategy-source'});`,
}));
assert.equal(strategyInput.ok, true);
assert.notEqual(strategyInput.result.path, strategySourcePath);

const smartArtCommitCell = [
  "const fs=await ctx.import('node:fs/promises');",
  "const path=await ctx.import('node:path');",
  "const {FileBlob,PresentationFile,reviewArtifact}=await ctx.import('office-kit');",
  `const source=await fs.readFile(${JSON.stringify(strategyInput.result.path)});`,
  "const presentation=await PresentationFile.importPptx(new FileBlob(source,{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'}));",
  "const targetId='presentation/slide/1/element/3';",
  "const leaves=presentation.inspect({includeNativeLeaves:true,target:targetId}).ndjson.split('\\n').filter(Boolean).map(JSON.parse);",
  "const leaf=leaves.find(item=>item.kind==='nativeLeaf'&&item.leafKind==='diagramText'&&item.value==='Scale candidate');",
  "if(!leaf) throw new Error('SmartArt leaf was not resolved from the reopened source');",
  "ctx.state.issuedLeafId=leaf.leafId;",
  "presentation.editNativeLeaf(targetId,leaf.leafId,{expectedHash:leaf.expectedHash,value:'Scale'});",
  "const output=await PresentationFile.exportPptx(presentation);",
  "const review=await reviewArtifact(output,{outputPath:path.join(ctx.taskRoot,'candidates','strategy-smartart.pptx'),layout:false,visualReview:'unavailable',verifyOptions:{minOverlapArea:46081}});",
  "const badPlan=structuredClone(output.metadata.editPlan);",
  "badPlan.operations[0].targetPartPath='../escape.xml';",
  "const badCandidate={metadata:{editPlan:badPlan},arrayBuffer:()=>output.arrayBuffer()};",
  "const rawPlan=structuredClone(output.metadata.editPlan);",
  "rawPlan.operations[0].rawXml='<a:t>unsafe</a:t>';",
  "const rawCandidate={metadata:{editPlan:rawPlan},arrayBuffer:()=>output.arrayBuffer()};",
  "const invalidCodes=[];",
  "for(const candidate of [badCandidate,rawCandidate]){try{await ctx.commit(candidate,{artifactId:'strategy-deck',kind:'presentation',name:'strategy-reviewed.pptx',summary:'Unsafe dependent binding',review});}catch(error){invalidCodes.push(error.code);}}",
  "const commit=await ctx.commit(output,{artifactId:'strategy-deck',kind:'presentation',name:'strategy-reviewed.pptx',summary:'Edit SmartArt scale label',review,next:'Reopen the reviewed revision and edit its detail title'});",
  "return {invalidCodes,selectedLeafId:leaf.leafId,commit};",
].join(" ");
const smartArtCommit = await strategySession.handleLine(JSON.stringify({ id: "edit-smartart", code: smartArtCommitCell }));
assert.equal(smartArtCommit.ok, true, JSON.stringify(smartArtCommit, null, 2));
assert.deepEqual(smartArtCommit.result.invalidCodes, ["invalid-edit-plan", "invalid-edit-plan"]);
assert.equal(smartArtCommit.result.commit.commitId, "c0001");
assert.deepEqual(smartArtCommit.result.commit.operation?.changedParts, ["ppt/diagrams/strategy-data.xml"]);
await strategySession.close();

const resumedStrategy = await createReplSession({ workspaceRoot: strategyWorkspace, taskId: strategyTaskId });
assert.equal(resumedStrategy.ready.resumedFrom.commitId, "c0001");
assert.equal(resumedStrategy.ready.operations.length, 1);
assert.equal(resumedStrategy.ready.operations[0].operationIds.length, 1);
assert.equal(resumedStrategy.ready.artifacts.length, 1);
const firstStrategyOperation = JSON.parse(await readFile(resumedStrategy.ready.operations[0].path, "utf8"));
assert.equal(firstStrategyOperation.plan.operations[0].leafKind, "diagramText");
assert.equal(firstStrategyOperation.plan.operations[0].targetPartPath, "ppt/diagrams/strategy-data.xml");
const firstReviewedPath = resumedStrategy.ready.artifacts[0].path;
const titleCommitCell = [
  "const fs=await ctx.import('node:fs/promises');",
  "const path=await ctx.import('node:path');",
  "const {FileBlob,PresentationFile,reviewArtifact}=await ctx.import('office-kit');",
  `const source=await fs.readFile(${JSON.stringify(firstReviewedPath)});`,
  "const presentation=await PresentationFile.importPptx(new FileBlob(source,{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'}));",
  "const smartArtTarget='presentation/slide/1/element/3';",
  "const smartArtLeaves=presentation.inspect({includeNativeLeaves:true,target:smartArtTarget}).ndjson.split('\\n').filter(Boolean).map(JSON.parse);",
  "const scaleLeaf=smartArtLeaves.find(item=>item.kind==='nativeLeaf'&&item.leafKind==='diagramText'&&item.value==='Scale');",
  "if(!scaleLeaf) throw new Error('Reviewed SmartArt edit was not restored');",
  "const titleTarget='presentation/slide/2/element/1';",
  "const titleLeaves=presentation.inspect({includeNativeLeaves:true,target:titleTarget}).ndjson.split('\\n').filter(Boolean).map(JSON.parse);",
  "const titleLeaf=titleLeaves.find(item=>item.kind==='nativeLeaf'&&item.leafKind==='text'&&item.value==='Strategy details');",
  "if(!titleLeaf) throw new Error('Title node index was not rebuilt from reviewed bytes');",
  "presentation.editNativeLeaf(titleTarget,titleLeaf.leafId,{expectedHash:titleLeaf.expectedHash,value:'Strategy evidence'});",
  "const output=await PresentationFile.exportPptx(presentation);",
  "const review=await reviewArtifact(output,{outputPath:path.join(ctx.taskRoot,'candidates','strategy-title.pptx'),layout:false,visualReview:'unavailable',verifyOptions:{minOverlapArea:46081}});",
  "const commit=await ctx.commit(output,{artifactId:'strategy-deck',kind:'presentation',name:'strategy-reviewed.pptx',summary:'Edit strategy detail title after resume',review,next:'Reopen, verify both edits, and publish'});",
  "return {heapRestored:ctx.state.issuedLeafId??null,observedScale:scaleLeaf.value,selectedLeafId:titleLeaf.leafId,commit};",
].join(" ");
const titleCommit = await resumedStrategy.handleLine(JSON.stringify({ id: "edit-title-after-resume", code: titleCommitCell }));
assert.equal(titleCommit.ok, true, JSON.stringify(titleCommit, null, 2));
assert.equal(titleCommit.result.heapRestored, null);
assert.equal(titleCommit.result.observedScale, "Scale");
assert.equal(titleCommit.result.commit.commitId, "c0002");
assert.notEqual(titleCommit.result.selectedLeafId, smartArtCommit.result.selectedLeafId);
await resumedStrategy.close();

const publishStrategy = await createReplSession({ workspaceRoot: strategyWorkspace, taskId: strategyTaskId });
assert.equal(publishStrategy.ready.resumedFrom.commitId, "c0002");
assert.equal(publishStrategy.ready.operations.length, 2);
assert.deepEqual(publishStrategy.ready.operations.map((operation) => operation.commitId), ["c0001", "c0002"]);
const secondStrategyOperation = JSON.parse(await readFile(publishStrategy.ready.operations[1].path, "utf8"));
assert.equal(secondStrategyOperation.plan.operations[0].leafKind, "text");
assert.deepEqual(secondStrategyOperation.plan.changedParts, ["ppt/slides/slide2.xml"]);
const finalReviewedPath = publishStrategy.ready.artifacts[0].path;
const finalVerification = await publishStrategy.handleLine(JSON.stringify({
  id: "verify-rebuilt-index",
  code: [
    "const fs=await ctx.import('node:fs/promises');",
    "const {FileBlob,PresentationFile}=await ctx.import('office-kit');",
    `const source=await fs.readFile(${JSON.stringify(finalReviewedPath)});`,
    "const presentation=await PresentationFile.importPptx(new FileBlob(source,{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'}));",
    "const records=['presentation/slide/1/element/3','presentation/slide/2/element/1'].flatMap(target=>presentation.inspect({includeNativeLeaves:true,target}).ndjson.split('\\n').filter(Boolean).map(JSON.parse));",
    "return {scale:records.some(item=>item.kind==='nativeLeaf'&&item.leafKind==='diagramText'&&item.value==='Scale'),title:records.some(item=>item.kind==='nativeLeaf'&&item.leafKind==='text'&&item.value==='Strategy evidence')};",
  ].join(" "),
}));
assert.deepEqual(finalVerification.result, { scale: true, title: true });
const publishedStrategy = await publishStrategy.handleLine(JSON.stringify({
  id: "publish-reviewed-strategy",
  code: "return await ctx.publish(ctx.task.commit,{artifactId:'strategy-deck',name:'strategy-reviewed.pptx'});",
}));
assert.equal(publishedStrategy.ok, true, JSON.stringify(publishedStrategy, null, 2));
assert.notEqual(publishedStrategy.result.path, strategySourcePath);
assert.deepEqual(await readFile(strategySourcePath), strategySourceBytes, "task workflow must not modify the source PPTX");
assert.deepEqual(await readFile(publishedStrategy.result.path), await readFile(finalReviewedPath));
await publishStrategy.close();

console.log("OfficeKit four-format task artifact smoke ok");
