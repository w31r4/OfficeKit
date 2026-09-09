import assert from "node:assert/strict";
import registry from "../src/ppj/capability-registry.json" with { type: "json" };
import schema from "../src/ppj/ppj-v1.schema.json" with { type: "json" };
import preview from "../src/ppj/svg-preview-capabilities.json" with { type: "json" };
import { readdir, readFile, stat } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { derivePreviewCapabilities, validatePreviewSupport, assertPreviewCapabilitiesCurrent, previewSchemaAt } from "../src/ppj/preview-capabilities.mjs";

assertPreviewCapabilitiesCurrent(preview, registry, schema);
assert.deepEqual(derivePreviewCapabilities(registry, schema), derivePreviewCapabilities(registry, schema));
const declared = [...preview.supported, ...preview.partial, ...preview.opaque, ...preview.unavailable];
assert.equal(new Set(declared).size, declared.length, "support grades must be disjoint");
assert.ok(!declared.includes("nativeRef") && !declared.includes("embeddedOle"));
assert.ok(declared.includes("ole"));
assert.equal(preview.variants.stream, "partial");
assert.equal(preview.sourceBound.nativeRef, "opaque");
assert.ok(!declared.includes("chart:streamgraph") && !declared.includes("chart:pictographic"));
for (const group of ["types", "charts", "fields", "variants", "sourceBound", "metadata", "factual"]) {
  for (const rule of Object.values(registry.previewSupport[group])) {
    assert.ok((await stat(rule.test)).isFile());
    if (rule.completeStateTest) assert.ok((await stat(rule.completeStateTest)).isFile());
    if (rule.stateTest) assert.ok((await stat(rule.stateTest)).isFile());
    assert.ok(previewSchemaAt(schema, rule.schemaRef));
  }
}

for (const mutate of [
  (r) => { delete r.previewSupport.types.ole; },
  (r) => { r.previewSupport.types.embeddedOle = r.previewSupport.types.ole; delete r.previewSupport.types.ole; },
  (r) => { r.previewSupport.charts.streamgraph = r.previewSupport.charts.area; },
  (r) => { r.previewSupport.types.shape.schemaRef = "#/$defs/imageElement"; },
  (r) => { r.previewSupport.fields.fill.schemaRef = "#/$defs/noSuchVisualField"; },
  (r) => { r.previewSupport.fields.fill.owner = "untracked-renderer"; },
  (r) => { r.previewSupport.charts.pie.status = "full"; },
  (r) => { r.previewSupport.charts.pie.status = "supported"; },
  (r) => { r.previewSupport.fields.text.test = ""; },
  (r) => { r.previewSupport.variants.pie = r.previewSupport.variants.symbol; },
  (r) => { r.previewSupport.metadata.meta.schemaRef = "#/$defs/notMetadata"; },
  (r) => { r.previewSupport.metadata.meta.status = "supported"; },
  (r) => { r.previewSupport.factual.chartMissing.severity = "warning"; },
]) {
  const changed = structuredClone(registry); mutate(changed);
  assert.throws(() => validatePreviewSupport(changed, schema));
}
const stale = structuredClone(preview); stale.supported.push("placeholder");
assert.throws(() => assertPreviewCapabilitiesCurrent(stale, registry, schema), /stale/);
const expandedSchema = structuredClone(schema);
expandedSchema.$defs.chartElement.allOf[1].properties.chartType.enum.push("future-chart");
assert.throws(() => validatePreviewSupport(registry, expandedSchema), /actual schema vocabulary/);

// --check verifies tracked state without regenerating it as a side effect.
const summaryPath = "src/ppj/svg-preview-capabilities.json";
const before = await readFile(summaryPath), modified = (await stat(summaryPath)).mtimeMs;
const check = spawnSync(process.execPath, ["scripts/generate-ppj-preview-capabilities.mjs", "--check"], { encoding: "utf8" });
assert.equal(check.status, 0, check.stderr);
assert.deepEqual(await readFile(summaryPath), before);
assert.equal((await stat(summaryPath)).mtimeMs, modified);

const fixtureRoot = path.resolve("test/fixtures");
async function walk(dir) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...await walk(file));
    else if (entry.name.endsWith(".ppj")) out.push(file);
  }
  return out;
}
function inspectTree(value, location, found = []) {
  if (!value || typeof value !== "object") return found;
  if (Object.hasOwn(value, "frame") && typeof value.type === "string") {
    assert.ok(Object.hasOwn(registry.previewSupport.types, value.type), `Missing preview element declaration: ${location}`);
    if (value.type === "chart" && value.chartType !== undefined) assert.ok(Object.hasOwn(registry.previewSupport.charts, value.chartType), `Missing preview chart declaration: ${location}`);
    found.push(value.type);
  }
  for (const [key, child] of Object.entries(value)) inspectTree(child, `${location}.${key}`, found);
  return found;
}
for (const file of await walk(fixtureRoot)) inspectTree(JSON.parse(await readFile(file, "utf8")), file);
assert.deepEqual(inspectTree({ pages: [{ elements: [{ type: "group", frame: {}, elements: [{ type: "image", frame: {} }] }] }] }, "$"), ["group", "image"]);
assert.throws(() => inspectTree({ type: "group", frame: {}, elements: [{ type: "unknown", frame: {} }] }, "$"), /Missing preview element/);
console.log(`ppj preview capability coverage ok (${Object.keys(registry.previewSupport.types).length} schema element types, ${Object.keys(registry.previewSupport.charts).length} chart types, registry-owned grades)`);
