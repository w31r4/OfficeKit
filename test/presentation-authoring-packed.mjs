import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { lstat, mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const temporary = await mkdtemp(path.join(os.tmpdir(), "officekit-presentation-packed-"));
const packDirectory = path.join(temporary, "pack");
const installDirectory = path.join(temporary, "installation");
const workspace = path.join(temporary, "workspace");

try {
  await mkdir(packDirectory, { recursive: true });
  await mkdir(installDirectory, { recursive: true });
  await mkdir(workspace, { recursive: true });
  const npm = process.platform === "win32" ? "npm.cmd" : "npm";
  const packed = run(npm, [
    "pack",
    "--ignore-scripts",
    "--json",
    "--pack-destination",
    packDirectory,
  ], { cwd: repoRoot });
  const packMetadata = JSON.parse(packed.stdout.trim())[0];
  assert.ok(packMetadata?.filename, "npm pack must return one tarball");
  const tarball = path.join(packDirectory, packMetadata.filename);

  run(npm, [
    "install",
    "--ignore-scripts",
    "--no-audit",
    "--no-fund",
    tarball,
  ], { cwd: installDirectory });
  const installedPackageRoot = path.join(installDirectory, "node_modules", "office-kit");
  const cli = path.join(installedPackageRoot, "bin", "officekit.mjs");
  const packageMetadata = JSON.parse(await readFile(path.join(installedPackageRoot, "package.json"), "utf8"));
  assert.equal(packageMetadata.name, "office-kit");
  assert.equal(packageMetadata.version, "0.6.0");
  assert.equal(await exists(path.join(workspace, "node_modules")), false, "task workspace must not have a local OfficeKit dependency");

  const init = run(process.execPath, [cli, "init", workspace, "--tools", "agents", "--json"], { cwd: workspace });
  const initResult = JSON.parse(init.stdout.trim());
  assert.equal(initResult.ok, true);
  const templateSearch = run(process.execPath, [
    cli,
    "template",
    "search",
    "--kind",
    "presentation",
    "--purpose",
    "quarterly business review",
    "--json",
  ], { cwd: workspace });
  const templateSearchResult = JSON.parse(templateSearch.stdout.trim());
  assert.equal(templateSearchResult.selectionMade, false);
  assert.ok(templateSearchResult.candidates.length > 0);
  assert.equal(await exists(path.join(workspace, ".office-kit", "providers")), false, "init and template search must not create provider state");

  const selfDirected = await runSelfDirected({ cli, workspace });
  assert.equal(selfDirected.published.reviewVerdict, "passed-with-limitations");
  assert.equal(path.basename(selfDirected.published.path), "self-directed.pptx");
  assert.equal(await exists(selfDirected.published.path), true);
  assert.match(selfDirected.published.sha256, /^[a-f0-9]{64}$/u);

  const templateConditioned = await runTemplateConditioned({ cli, workspace: path.join(temporary, "template-workspace") });
  assert.equal(templateConditioned.published.reviewVerdict, "passed-with-limitations");
  assert.equal(path.basename(templateConditioned.published.path), "template-conditioned.pptx");
  assert.equal(templateConditioned.evidence.designProfile, "office-kit/pptx-design-profile/v1");
  assert.equal(templateConditioned.evidence.templatePlan, "ready");
  assert.equal(templateConditioned.evidence.sourceReuse, true);
  assert.match(templateConditioned.published.sha256, /^[a-f0-9]{64}$/u);

  console.log("presentation authoring packed smoke ok");
} finally {
  await rm(temporary, { recursive: true, force: true });
}

async function runSelfDirected({ cli, workspace: root }) {
  const taskRoot = path.resolve(root);
  const firstCell = path.join(taskRoot, "self-directed-1.mjs");
  const secondCell = path.join(taskRoot, "self-directed-2.mjs");
  const thirdCell = path.join(taskRoot, "self-directed-3.mjs");
  await writeFile(firstCell, selfDirectedFirstCell());
  await writeFile(secondCell, selfDirectedSecondCell());
  await writeFile(thirdCell, selfDirectedThirdCell());
  assertCellIsTyped(firstCell);
  assertCellIsTyped(secondCell);
  assertCellIsTyped(thirdCell);

  const first = parseJsonLines(run(process.execPath, [
    cli,
    "repl",
    "--new",
    "Create a concise architecture decision deck",
    "--workspace",
    taskRoot,
    "--file",
    firstCell,
  ], { cwd: taskRoot }).stdout);
  assert.equal(first[0].type, "session.ready");
  assert.equal(first[1].ok, true, JSON.stringify(first[1], null, 2));
  const taskId = first[0].task.id;
  assert.equal(first[1].result.plan.mode, "create");
  assert.equal(first[1].result.commit.commitId, "c0001");

  const second = parseJsonLines(run(process.execPath, [
    cli,
    "repl",
    taskId,
    "--workspace",
    taskRoot,
    "--file",
    secondCell,
  ], { cwd: taskRoot }).stdout);
  assert.equal(second[0].type, "session.ready");
  assert.equal(second[0].resumedFrom.commitId, "c0001");
  assert.equal(second[1].ok, true, JSON.stringify(second[1], null, 2));
  assert.equal(second[1].result.commit.commitId, "c0002");
  assert.equal(second[1].result.changedText, "Architecture decision — reviewed");

  const third = parseJsonLines(run(process.execPath, [
    cli,
    "repl",
    taskId,
    "--workspace",
    taskRoot,
    "--file",
    thirdCell,
  ], { cwd: taskRoot }).stdout);
  assert.equal(third[0].type, "session.ready");
  assert.equal(third[0].resumedFrom.commitId, "c0002");
  assert.equal(third[1].ok, true, JSON.stringify(third[1], null, 2));
  assert.equal(path.basename(third[1].result.published.path), "self-directed.pptx");
  return third[1].result;
}

async function runTemplateConditioned({ cli, workspace: root }) {
  const taskRoot = path.resolve(root);
  await mkdir(taskRoot, { recursive: true });
  const firstCell = path.join(taskRoot, "template-1.mjs");
  const secondCell = path.join(taskRoot, "template-2.mjs");
  const thirdCell = path.join(taskRoot, "template-3.mjs");
  await writeFile(firstCell, templateFirstCell());
  await writeFile(secondCell, templateSecondCell());
  await writeFile(thirdCell, templateThirdCell());
  assertCellIsTyped(firstCell);
  assertCellIsTyped(secondCell);
  assertCellIsTyped(thirdCell);

  const first = parseJsonLines(run(process.execPath, [
    cli,
    "repl",
    "--new",
    "Create a new deck from an authoritative template",
    "--workspace",
    taskRoot,
    "--file",
    firstCell,
  ], { cwd: taskRoot }).stdout);
  assert.equal(first[0].type, "session.ready");
  assert.equal(first[1].ok, true, JSON.stringify(first[1], null, 2));
  const taskId = first[0].task.id;
  assert.equal(first[1].result.evidence.templatePlan, "ready");
  assert.equal(first[1].result.evidence.sourceReuse, true);
  assert.equal(first[1].result.commit.commitId, "c0001");

  const second = parseJsonLines(run(process.execPath, [
    cli,
    "repl",
    taskId,
    "--workspace",
    taskRoot,
    "--file",
    secondCell,
  ], { cwd: taskRoot }).stdout);
  assert.equal(second[0].resumedFrom.commitId, "c0001");
  assert.equal(second[1].ok, true, JSON.stringify(second[1], null, 2));
  assert.equal(second[1].result.commit.commitId, "c0002");
  assert.equal(second[1].result.changedText, "Template-derived evidence — reviewed");

  const third = parseJsonLines(run(process.execPath, [
    cli,
    "repl",
    taskId,
    "--workspace",
    taskRoot,
    "--file",
    thirdCell,
  ], { cwd: taskRoot }).stdout);
  assert.equal(third[0].resumedFrom.commitId, "c0002");
  assert.equal(third[1].ok, true, JSON.stringify(third[1], null, 2));
  assert.equal(path.basename(third[1].result.published.path), "template-conditioned.pptx");
  return { ...third[1].result, evidence: first[1].result.evidence };
}

function selfDirectedFirstCell() {
  return [
    'const { Presentation, PresentationFile, reviewArtifact } = await ctx.import("office-kit");',
    'const plan = { schema:"office-kit/presentation-authoring-plan/v1", mode:"create", brief:{audience:"Engineering leadership", purpose:"Choose a migration path", expectedOutcome:"One bounded decision"}, narrative:{thesis:"The bounded path is the safer decision", sections:["Context","Decision"]}, design:{sourceMode:"self-directed", mechanismPacks:["technical-architecture"], designGrammar:{density:"one claim per page", rhythm:"sparse conclusion then evidence"}}, pages:[{id:"p01-decision", readerTask:"Understand the recommendation", claim:"Choose the bounded path", evidence:["Measured build time"], contentBudget:{maxCharacters:240,maxObjects:8}, compositionIntent:"One dominant conclusion and one evidence line"}], editorial:{voice:"direct and evidence-led", lockedFacts:["measured build time"], avoid:["empty transition phrases"]}, artifactRefs:[], recipe:"tasks/create.md", unresolved:[], nextAction:"Review and continue the deck"};',
    'await ctx.plan(plan);',
    'const presentation = Presentation.create({ slideSize:{width:960,height:540} });',
    'const slide = presentation.slides.add({ name:"Decision" });',
    'slide.shapes.add({ name:"Decision headline", geometry:"textbox", text:[{runs:[{text:"Architecture decision",style:{bold:true,fontSize:28,fontFamily:"Arial"}}]}], position:{left:60,top:60,width:760,height:80} });',
    'slide.shapes.add({ name:"Evidence", geometry:"textbox", text:"Use the bounded path", position:{left:60,top:180,width:760,height:100} });',
    'const candidate = await PresentationFile.exportPptx(presentation);',
    'const review = await reviewArtifact(candidate,{authoringPlan:plan,changedPageIds:["p01-decision"],outputPath:`${ctx.taskRoot}/evidence/self-directed-review.pptx`,layout:false,visualReview:"unavailable"});',
    'const commit = await ctx.commit(candidate,{artifactId:"deck",kind:"presentation",name:"self-directed-working.pptx",summary:"Create the first plan-bound working deck",review,next:"Resume and refine the decision headline"});',
    'return {plan:await ctx.plan(),commit};',
  ].join("\n") + "\n";
}

function selfDirectedSecondCell() {
  return [
    'const fs = await ctx.import("node:fs/promises");',
    'const path = await ctx.import("node:path");',
    'const { FileBlob, PresentationFile, reviewArtifact } = await ctx.import("office-kit");',
    'const artifact = ctx.task.artifacts.find((item) => item.id === "deck");',
    'const bytes = await fs.readFile(path.resolve(ctx.taskRoot, artifact.headRevision.path));',
    'const presentation = await PresentationFile.importPptx(new FileBlob(bytes,{type:"application/vnd.openxmlformats-officedocument.presentationml.presentation"}));',
    'presentation.slides.getItem(0).shapes.getItemAt(0).text.set("Architecture decision — reviewed");',
    'const candidate = await PresentationFile.exportPptx(presentation);',
    'const plan = await ctx.plan();',
    'const review = await reviewArtifact(candidate,{authoringPlan:plan,changedPageIds:["p01-decision"],outputPath:`${ctx.taskRoot}/evidence/self-directed-review-2.pptx`,layout:false,visualReview:"unavailable"});',
    'const commit = await ctx.commit(candidate,{artifactId:"deck",kind:"presentation",name:"self-directed-working-v2.pptx",summary:"Refine the decision headline after resume",review,next:"Publish the reviewed deck"});',
    'return {changedText:presentation.slides.getItem(0).shapes.getItemAt(0).text.value,commit};',
  ].join("\n") + "\n";
}

function selfDirectedThirdCell() {
  return [
    'const publication = await ctx.publish(ctx.task.commit,{artifactId:"deck",name:"self-directed.pptx"});',
    'return {published:publication};',
  ].join("\n") + "\n";
}

function templateFirstCell() {
  return [
    'const path = await ctx.import("node:path");',
    'const { Presentation, PresentationFile, FileBlob, reviewArtifact } = await ctx.import("office-kit");',
    'const template = Presentation.create({ slideSize:{width:960,height:540} });',
    'for (const [index, text] of ["Source title one","Source title two"].entries()) { const slide = template.slides.add({name:`Template ${index+1}`}); slide.shapes.add({name:"Template title",geometry:"textbox",text:[{runs:[{text,style:{bold:true,fontSize:26,fontFamily:"Arial"}}]}],position:{left:60,top:60,width:760,height:90}}); }',
    'const templateFile = await PresentationFile.exportPptx(template);',
    'const templatePath = path.join(ctx.workspaceRoot,"template-source.pptx");',
    'await templateFile.save(templatePath);',
    'const staged = await ctx.input(templatePath,{artifactId:"authoritative-template",kind:"presentation"});',
    'const plan = {schema:"office-kit/presentation-authoring-plan/v1",mode:"create-from-template",brief:{audience:"Product reviewers",purpose:"Turn a supplied template into a decision deck",expectedOutcome:"A template-consistent draft"},narrative:{thesis:"Evidence should inherit the supplied visual language",sections:["Evidence","Decision"]},design:{sourceMode:"template",mechanismPacks:["enterprise-data-review"],designGrammar:{sourceArtifact:"authoritative-template",density:"reuse the source frame"},artifactRef:{artifactId:staged.artifactId,sha256:staged.sha256}},pages:[{id:"p01-source",readerTask:"Read the first source frame",claim:"The first source frame remains unchanged",evidence:["Template title frame"],contentBudget:{maxCharacters:240,maxObjects:8},compositionIntent:"Preserve the source frame"},{id:"p02-source",readerTask:"Read the second source frame",claim:"The second source frame remains unchanged",evidence:["Template title frame"],contentBudget:{maxCharacters:240,maxObjects:8},compositionIntent:"Preserve the source frame"},{id:"p03-evidence",readerTask:"Read the generated evidence",claim:"The source frame is reused safely",evidence:["Template title frame"],contentBudget:{maxCharacters:240,maxObjects:8},compositionIntent:"Reuse the source title component and replace only its text"}],editorial:{voice:"plain and evidence-led",lockedFacts:["source frame"],avoid:["generic template language"]},artifactRefs:[{artifactId:staged.artifactId,sha256:staged.sha256,role:"authoritative-template"}],recipe:"tasks/create-from-template.md",unresolved:[],nextAction:"Review the source-derived page"};',
    'await ctx.plan(plan);',
    'const sourceBytes = await fsRead(staged.path);',
    'const imported = await PresentationFile.importPptx(new FileBlob(sourceBytes,{type:"application/vnd.openxmlformats-officedocument.presentationml.presentation"}));',
    'const profile = imported.designProfile({maxItems:32});',
    'const templatePlan = imported.planTemplateGeneration({slides:[{role:"title",title:"New evidence"}]});',
    'if (templatePlan.status !== "ready") throw new Error(`template plan was ${templatePlan.status}`);',
    'const page = templatePlan.pages[0];',
    'imported.reuseSourceSlide({slideId:page.sourceSlideId,sourceRevisionSha256:templatePlan.source.revisionSha256,expectedCloneCapability:page.source.cloneCapability});',
    'const clonedBytes = await (await PresentationFile.exportPptx(imported)).arrayBuffer();',
    'const continued = await PresentationFile.importPptx(new FileBlob(new Uint8Array(clonedBytes),{type:"application/vnd.openxmlformats-officedocument.presentationml.presentation"}));',
    'continued.slides.items.at(-1).shapes.getItemAt(0).text.set("Template-derived evidence");',
    'const candidate = await PresentationFile.exportPptx(continued);',
    'const review = await reviewArtifact(candidate,{authoringPlan:plan,changedPageIds:["p03-evidence"],outputPath:path.join(ctx.taskRoot,"evidence/template-review.pptx"),layout:false,visualReview:"unavailable"});',
    'const commit = await ctx.commit(candidate,{artifactId:"deck",kind:"presentation",name:"template-working.pptx",summary:"Create a source-derived template draft",review,next:"Resume and refine the reused frame"});',
    'return {evidence:{designProfile:profile.schema,templatePlan:templatePlan.status,sourceReuse:true},commit};',
    'async function fsRead(file){ const fs = await ctx.import("node:fs/promises"); return fs.readFile(file); }',
  ].join("\n") + "\n";
}

function templateSecondCell() {
  return [
    'const fs = await ctx.import("node:fs/promises");',
    'const path = await ctx.import("node:path");',
    'const { FileBlob, PresentationFile, reviewArtifact } = await ctx.import("office-kit");',
    'const artifact = ctx.task.artifacts.find((item) => item.id === "deck");',
    'const bytes = await fs.readFile(path.resolve(ctx.taskRoot, artifact.headRevision.path));',
    'const presentation = await PresentationFile.importPptx(new FileBlob(bytes,{type:"application/vnd.openxmlformats-officedocument.presentationml.presentation"}));',
    'presentation.slides.getItem(2).shapes.getItemAt(0).text.set("Template-derived evidence — reviewed");',
    'const candidate = await PresentationFile.exportPptx(presentation);',
    'const plan = await ctx.plan();',
    'const review = await reviewArtifact(candidate,{authoringPlan:plan,changedPageIds:["p03-evidence"],outputPath:`${ctx.taskRoot}/evidence/template-review-2.pptx`,layout:false,visualReview:"unavailable"});',
    'const commit = await ctx.commit(candidate,{artifactId:"deck",kind:"presentation",name:"template-working-v2.pptx",summary:"Refine the source-derived page after resume",review,next:"Publish the template-conditioned deck"});',
    'return {changedText:presentation.slides.getItem(2).shapes.getItemAt(0).text.value,commit};',
  ].join("\n") + "\n";
}

function templateThirdCell() {
  return [
    'const publication = await ctx.publish(ctx.task.commit,{artifactId:"deck",name:"template-conditioned.pptx"});',
    'return {published:publication};',
  ].join("\n") + "\n";
}

function assertCellIsTyped(file) {
  const source = requireText(file);
  assert.doesNotMatch(source, /(?:JSZip|raw\s+OOXML|xpath|@oai\/artifact-tool|python|PPTD)/iu, `${file} must use the public typed API`);
}

function requireText(file) {
  return readFileSync(file, "utf8");
}

function parseJsonLines(source) {
  return source.trim().split(/\r?\n/u).filter(Boolean).map((line) => JSON.parse(line));
}

function run(command, args, { cwd }) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    env: { ...process.env, npm_config_fund: "false", npm_config_audit: "false" },
    shell: false,
  });
  assert.equal(result.status, 0, `${command} ${args.join(" ")} failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  return result;
}

async function exists(target) {
  try {
    await lstat(target);
    return true;
  } catch (error) {
    if (error.code === "EISDIR") return true;
    if (error.code === "ENOENT") return false;
    throw error;
  }
}
