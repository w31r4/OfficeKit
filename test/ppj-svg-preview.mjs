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

const taskRoot = await mkdtemp(path.join(os.tmpdir(), "officekit-ppj-preview-"));
const out = path.join(taskRoot, "preview");
for (const type of ["shape", "text", "image", "table", "connector", "group"]) assert.ok(SVG_PREVIEW_SUPPORTED_TYPES.has(type));
assert.ok(previewCapabilities.supported.includes("chart:combo"));
const result = await renderPpjToSvg("test/fixtures/presentation/evidence-ledger-canonical.ppj", { outputDir: out });
assert.equal(result.renderer, "officekit-svg-preview");
assert.equal(result.pages.length, 2);
assert.ok((await readFile(path.join(out, "page-claim.svg"), "utf8")).includes("data-officekit-id=\"claim-title\""));
assert.ok((await readFile(path.join(out, "page-claim.png"))).byteLength > 1000);
assert.equal(JSON.parse(await readFile(path.join(out, "render.json"), "utf8")).pages.length, 2);
assert.equal(result.receipt.output.status, "complete");
assert.equal(result.receipt.environment.raster.name, "sharp");
assert.ok(result.receipt.environment.raster.versions.sharp);
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
