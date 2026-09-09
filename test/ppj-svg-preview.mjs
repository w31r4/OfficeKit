import assert from "node:assert/strict";
import { readFile, writeFile, mkdtemp } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { renderPpjToSvg, SVG_PREVIEW_SUPPORTED_TYPES } from "../src/ppj/svg-preview.mjs";
import previewCapabilities from "../src/ppj/svg-preview-capabilities.json" with { type: "json" };
import JSZip from "jszip";
import { loadPpjWorkspace, compilePpjWorkspace, sha256 } from "../src/ppj/workspace.mjs";
import { projectPptxToPpj } from "../src/ppj/native.mjs";
import { spawnSync } from "node:child_process";

async function verifyAssessment(result) {
  const persisted = JSON.parse(await readFile(path.join(result.receipt.output.directory, "render.json"), "utf8"));
  assert.deepEqual(persisted, result.receipt);
  assert.deepEqual(result.assessment, persisted.assessment);
  assert.deepEqual(result.reliability, persisted.reliability);
  assert.equal(result.status, persisted.status);
  const { default: sharp } = await import("sharp");
  for (const [index, page] of result.pages.entries()) {
    const record = persisted.pages[index];
    assert.deepEqual(page.assessment, record.assessment);
    assert.deepEqual(page.diagnostics, record.diagnostics);
    assert.deepEqual(page.reliability, record.reliability);
    assert.equal(page.status, record.status);
    const svg = await readFile(path.join(result.receipt.output.directory, record.file), "utf8");
    assert.equal(svg, page.svg);
    const png = await readFile(path.join(result.receipt.output.directory, record.png));
    const metadata = await sharp(png).metadata();
    assert.equal(metadata.width, result.canvas.width);
    assert.equal(metadata.height, result.canvas.height);
    assert.notEqual(page.reliability.status, "passed", "these fixtures still have known visual limitations");
    assert.ok(svg.includes(`data-officekit-review="${page.reliability.status}"`));
    const pixel = await sharp(png).extract({ left: 1, top: 1, width: 1, height: 1 }).removeAlpha().raw().toBuffer();
    assert.deepEqual([...pixel], page.reliability.status === "failed" ? [153, 27, 27] : [146, 64, 14], "actual PNG must contain the warning banner");
  }
}

const taskRoot = await mkdtemp(path.join(os.tmpdir(), "officekit-ppj-preview-"));
const out = path.join(taskRoot, "preview");
// Whole-family support is not inferred from the existence of a draw branch.
assert.deepEqual([...SVG_PREVIEW_SUPPORTED_TYPES], previewCapabilities.supported.filter((type) => !type.includes(":")));
for (const type of ["shape", "text", "image", "table", "connector", "group", "chart:combo"]) assert.ok(previewCapabilities.partial.includes(type));
const result = await renderPpjToSvg("test/fixtures/presentation/evidence-ledger-canonical.ppj", { outputDir: out });
assert.equal(result.renderer, "officekit-svg-preview");
assert.equal(result.pages.length, 2);
assert.ok((await readFile(path.join(out, "page-claim.svg"), "utf8")).includes("data-officekit-id=\"claim-title\""));
assert.ok((await readFile(path.join(out, "page-claim.png"))).byteLength > 1000);
assert.equal(JSON.parse(await readFile(path.join(out, "render.json"), "utf8")).pages.length, 2);
assert.equal(result.receipt.output.status, "complete");
assert.equal(result.receipt.environment.raster.name, "sharp");
assert.ok(result.receipt.environment.raster.versions.sharp);
assert.equal(result.reliability.status, "failed", "known authored geometry/topology errors cannot pass by publishing");
assert.equal(result.ok, true, "publication completion remains separate from reliability");
await verifyAssessment(result);
const cli = spawnSync(process.execPath, ["bin/officekit.mjs", "ppj", "preview",
  "test/fixtures/presentation/evidence-ledger-canonical.ppj", "-o", path.join(taskRoot, "cli-preview"), "--json"], { encoding: "utf8", maxBuffer: 16 * 1024 * 1024 });
assert.equal(cli.status, 0, cli.stderr);
const cliResult = JSON.parse(cli.stdout);
assert.equal(cliResult.ok, true);
assert.equal(cliResult.reliability.status, "failed", "CLI exit zero is publication success only");
for (const artifact of result.receipt.artifacts) {
  const data = await readFile(path.join(out, artifact.file));
  assert.equal(artifact.sha256, sha256(data));
  assert.equal(artifact.bytes, data.byteLength);
}

// Reproject native parts rather than restoring the private authored snapshot,
// matching the existing source-bound test pattern in office-kit-skill.mjs.
const authored = await compilePpjWorkspace(await loadPpjWorkspace("examples/ppj/minimum.ppj"));
const zip = await JSZip.loadAsync(authored.file);
for (const name of Object.keys(zip.files)) if (name.startsWith("officeKit/")) zip.remove(name);
zip.file("_rels/.rels", (await zip.file("_rels/.rels").async("string")).replace(
  /<Relationship\b(?=[^>]*\bType="https:\/\/schemas\.officekit\.dev\/relationships\/presentation-program")[^>]*(?:\/>|>[\s\S]*?<\/Relationship>)/g, "",
));
zip.file("[Content_Types].xml", (await zip.file("[Content_Types].xml").async("string")).replace(
  /<Override\b(?=[^>]*\bPartName="\/officeKit\/)[^>]*(?:\/>|>[\s\S]*?<\/Override>)/g, "",
));
const source = await zip.generateAsync({ type: "uint8array" });
const projected = await projectPptxToPpj(source, { sourceUri: "source.pptx", assetRootUri: "assets" });
assert.equal(projected.sourceBound, true);
assert.equal(projected.assets.length, 0);
const sourcePath = path.join(taskRoot, "source.pptx");
const ppjPath = path.join(taskRoot, "imported.ppj");
await writeFile(sourcePath, source, { flag: "wx" });
await writeFile(ppjPath, projected.programJson, { flag: "wx" });
const imported = await renderPpjToSvg(ppjPath, { outputDir: path.join(taskRoot, "source-preview") });
assert.equal(imported.receipt.sourceBound, true);
assert.equal(imported.receipt.source.sha256, sha256(source));
assert.equal(imported.receipt.input.sha256, sha256(projected.programJson));
assert.equal(sha256(await readFile(sourcePath)), sha256(source));
assert.equal(sha256(await readFile(ppjPath)), sha256(projected.programJson));
assert.equal(imported.receipt.compile.outputSha256, sha256((await compilePpjWorkspace(await loadPpjWorkspace(ppjPath))).file));
assert.equal(imported.receipt.output.status, "complete");
assert.ok(imported.diagnostics.some((diagnostic) => diagnostic.reason.startsWith("preview.source")));
await verifyAssessment(imported);
for (const artifact of imported.receipt.artifacts) {
  const bytes = await readFile(path.join(imported.receipt.output.directory, artifact.file));
  assert.equal(artifact.sha256, sha256(bytes));
  assert.equal(artifact.bytes, bytes.byteLength);
}
const repeated = spawnSync(process.execPath, ["bin/officekit.mjs", "ppj", "preview", ppjPath, "-o", imported.receipt.output.directory, "--json"], { encoding: "utf8" });
assert.equal(repeated.status, 1, repeated.stderr);
assert.match(repeated.stderr, /preview\.output\.exists/);
assert.equal(sha256(await readFile(sourcePath)), sha256(source));
console.log(`ppj svg preview ok (authored and source-bound evidence: ${taskRoot})`);
