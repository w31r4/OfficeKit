import assert from "node:assert/strict";
import { readFile, mkdtemp } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { renderPpjToSvg, SVG_PREVIEW_SUPPORTED_TYPES } from "../src/ppj/svg-preview.mjs";
import previewCapabilities from "../src/ppj/svg-preview-capabilities.json" with { type: "json" };

const out = await mkdtemp(path.join(os.tmpdir(), "officekit-ppj-preview-"));
for (const type of ["shape", "text", "image", "table", "connector", "group"]) assert.ok(SVG_PREVIEW_SUPPORTED_TYPES.has(type));
assert.ok(previewCapabilities.supported.includes("chart:combo"));
const result = await renderPpjToSvg("test/fixtures/presentation/evidence-ledger-canonical.ppj", { outputDir: out });
assert.equal(result.renderer, "officekit-svg-preview");
assert.equal(result.pages.length, 2);
assert.ok((await readFile(path.join(out, "page-claim.svg"), "utf8")).includes("data-officekit-id=\"claim-title\""));
assert.ok((await readFile(path.join(out, "page-claim.png"))).byteLength > 1000);
assert.equal(JSON.parse(await readFile(path.join(out, "render.json"), "utf8")).pages.length, 2);
console.log("ppj svg preview ok");
