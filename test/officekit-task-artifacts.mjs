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

console.log("OfficeKit four-format task artifact smoke ok");
