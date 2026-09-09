import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { derivePreviewCapabilities } from "../src/ppj/preview-capabilities.mjs";

const root = path.resolve(import.meta.dirname, "..");
const registry = JSON.parse(await readFile(path.join(root, "src/ppj/capability-registry.json"), "utf8"));
const schema = JSON.parse(await readFile(path.join(root, "src/ppj/ppj-v1.schema.json"), "utf8"));
const previewCapabilities = derivePreviewCapabilities(registry, schema);
const boundaries = registry.authoredCompilerBoundaries.map((entry) => ({
  feature: entry.feature,
  ppjPath: entry.ppjPath,
  behavior: entry.behavior,
  sourceBound: entry.sourceBound,
  reason: entry.reason,
  visualStatus: entry.behavior === "full" ? "unreviewed" : "partial-or-opaque",
}));
const matrix = {
  schema: "office-kit/presentation-capability-matrix/v1",
  sources: {
    registry: "src/ppj/capability-registry.json",
    languageSchema: "src/ppj/ppj-v1.schema.json",
    generatedReference: registry.generatedReference,
    semanticValidator: "native/OfficeKit/src/OfficeKit.Codec/PpjSemanticValidator.cs",
    projector: "native/OfficeKit/src/OfficeKit.Codec/PpjPresentationProjector.cs",
    authoredCompiler: "native/OfficeKit/src/OfficeKit.Codec/PpjAuthoredPresentationCompiler.cs",
    chartCompiler: "native/OfficeKit/src/OfficeKit.Codec/PpjAuthoredChartCompiler.cs",
    previewCapabilities: "src/ppj/svg-preview-capabilities.json",
  },
  counts: {
    schemaDefinitions: Object.keys(schema.$defs || {}).length,
    ppjRootPaths: Object.keys(registry.ppjPathOwners).length,
    ppjStateApiPaths: Object.keys(registry.ppjStateApiPaths).length,
    authoredCompilerBoundaries: boundaries.length,
    nativeLeafKinds: Object.keys(registry.nativeLeafKinds).length,
    helpApis: Object.keys(registry.helpApis).length,
    hostOnlyApis: registry.hostOnly.length,
  },
  surfaces: {
    ppjState: "schema + semantic validator + authored compiler + projector",
    nativeRef: "imported source proof + closed native leaf vocabulary",
    inspectReview: "CLI inspect/check/render/review and review evidence",
    visualPreview: "must consume validated projection and report supported/partial/opaque/unavailable",
    exportFidelity: "PPTX build + re-import + source/stable-id verification where applicable",
  },
  rendererReuse: {
    input: "compilePpjWorkspace(...).programJson",
    semanticOwner: "native/OfficeKit/src/OfficeKit.Codec/PpjPresentationProjector.cs",
    layoutOwner: "native/OfficeKit/src/OfficeKit.Codec/PpjPresentationCompiler.cs",
    rasterBackend: "sharp (SVG to PNG)",
    officeCompatibilityBackend: "optional src/renderers/libreoffice.mjs",
    driftRule: "Do not add renderer-only capability lists; update registry, compiler/projector, preview capability map, fixture, and test together.",
  },
  previewCapabilities,
  previewSupport: registry.previewSupport,
  ppjRootPaths: Object.entries(registry.ppjPathOwners).map(([path, value]) => ({ path, ...value })),
  authoredCompilerBoundaries: boundaries,
  nativeLeafKinds: Object.entries(registry.nativeLeafKinds).map(([name, value]) => ({ name, ...value })),
  helpApis: Object.entries(registry.helpApis).map(([name, surface]) => ({ name, surface })),
  hostOnly: registry.hostOnly,
};
const args = process.argv.slice(2), check = args.includes("--check");
const destinations = args.filter((arg) => arg !== "--check");
if (destinations.length > 1 || destinations.some((arg) => arg.startsWith("--"))) throw new Error("Usage: generate-presentation-capability-matrix.mjs [output] [--check]");
const output = destinations[0] || path.join(root, "docs/presentation-capability-matrix.json");
const content = `${JSON.stringify(matrix, null, 2)}\n`;
if (check) {
  if (await readFile(output, "utf8") !== content) throw new Error("Presentation capability matrix is stale; run npm run docs:presentation-capabilities");
  console.log("presentation capability matrix is current");
} else {
  await writeFile(output, content);
  console.log(`wrote ${path.relative(root, output)} (${matrix.counts.authoredCompilerBoundaries} authored boundaries, ${matrix.counts.nativeLeafKinds} native leaves)`);
}
