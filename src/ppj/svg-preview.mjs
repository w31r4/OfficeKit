import path from "node:path";
import { createRequire } from "node:module";
import { compilePpjWorkspace, loadPpjWorkspace } from "./workspace.mjs";
import { previewInputEvidence, publishPpjPreview } from "./preview-output.mjs";
import { renderPpjSceneSvg } from "./preview-scene-svg.mjs";

const require = createRequire(import.meta.url);
const PREVIEW_CAPABILITIES = require("./svg-preview-capabilities.json");
// Whole-family coverage stays registry-owned even when more native fields paint.
export const SVG_PREVIEW_SUPPORTED_TYPES = new Set(PREVIEW_CAPABILITIES.supported.filter((value) => !value.includes(":")));

export async function renderPpjToSvg(inputPath, {
  cwd = process.cwd(), outputDir,
  load = loadPpjWorkspace, compile = compilePpjWorkspace, loadRaster,
} = {}) {
  const absolute = path.resolve(cwd, inputPath);
  const workspace = await load(absolute, { cwd, retainRoot: true });
  const compiled = await compile(workspace, { includeNodeMap: false, includePreviewScene: true });
  // Scene/candidate/asset validation precedes painting and publication. Missing
  // or incompatible native state never falls back to interpreting canonical PPJ.
  // Input assessment is mandatory here; ordinary build/check remain scene-free.
  const result = renderPpjSceneSvg(compiled);
  // Do not expose protobuf BigInts/asset graphs through the JSON-facing CLI.
  // The publisher receives the validated native scene; callers receive identity.
  const publicResult = { ...result, scene: result.sceneEvidence };
  if (outputDir) {
    const { receipt, warnings } = await publishPpjPreview(result, previewInputEvidence(workspace, compiled), { cwd, outputDir, loadRaster });
    return { ...publicResult, ...receipt, pages: result.pages, receipt, warnings };
  }
  return publicResult;
}
