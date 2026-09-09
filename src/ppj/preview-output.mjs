import { mkdir, readFile, unlink } from "node:fs/promises";
import path from "node:path";
import { sha256, writeExclusiveFile } from "./workspace.mjs";
import { previewAssessment, previewDiagnostic } from "./preview-diagnostics.mjs";

const SCHEMA = "office-kit/ppj-svg-preview-output/v1";
const jsonBytes = (value) => Buffer.from(`${JSON.stringify(value, null, 2)}\n`);

export class PreviewOutputError extends Error {
  constructor(message, code, receipt, cause) {
    super(message, { cause });
    this.name = "PreviewOutputError";
    this.code = code;
    this.outputPath = receipt.output.directory;
    this.receipt = receipt;
  }
}

// Keep filenames independent of arbitrary PPJ identifiers, including Windows
// device names and collisions on case-insensitive filesystems.
export function previewPageStems(pages) {
  const used = new Set();
  return pages.map(({ id }, index) => {
    const safe = typeof id === "string" && /^[a-z0-9][a-z0-9_-]{0,119}$/iu.test(id)
      && !/^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/iu.test(id);
    let stem = safe ? id : `slide-${index + 1}`;
    let suffix = 0;
    while (used.has(stem.toLowerCase())) stem = `slide-${index + 1}-${++suffix}`;
    used.add(stem.toLowerCase());
    return stem;
  });
}

export function previewInputEvidence(workspace, compiled) {
  return {
    input: { path: workspace.path, sha256: sha256(workspace.program) },
    compile: {
      programSha256: compiled.programSha256,
      outputSha256: sha256(compiled.file),
    },
    sourceBound: compiled.sourceBound === true,
    source: workspace.source?.byteLength
      ? { path: workspace.sourcePath, sha256: sha256(workspace.source) } : null,
    assets: workspace.assets.map((asset) => ({
      id: asset.id, sha256: sha256(asset.data), bytes: asset.data.byteLength,
    })),
  };
}

async function implementationIdentity(renderer) {
  const files = ["svg-preview.mjs", "preview-output.mjs", "svg-preview-capabilities.json",
    "preview-diagnostics.mjs", "preview-input-assessment.mjs", "preview-factual-errors.mjs",
    "preview-capabilities.mjs", "capability-registry.json", "ppj-v1.schema.json"];
  if (renderer === "officekit-native-scene-svg-internal") files.push("preview-scene-svg.mjs", "preview-scene-view.mjs", "preview-scene.mjs", "preset-geometry-profiles.json");
  return {
    version: 2,
    sources: await Promise.all(files.map(async (file) => ({
      file, sha256: sha256(await readFile(new URL(file, import.meta.url))),
    }))),
  };
}

async function loadSharpRaster() {
  const mod = await import("sharp");
  const sharp = mod.default || mod;
  return {
    identity: { name: "sharp", status: "available", versions: { ...sharp.versions } },
    render: (svg) => sharp(Buffer.from(svg)).png().toBuffer(),
  };
}

// Injection is confined to this leaf publisher so failure cases do not require
// installing/removing dependencies or making the machine's filesystem fail.
export async function publishPpjPreview(result, evidence, {
  outputDir,
  cwd = process.cwd(),
  loadRaster = loadSharpRaster,
  writeArtifact = writeExclusiveFile,
  removePending = unlink,
} = {}) {
  const destination = path.resolve(cwd, outputDir);
  const stems = previewPageStems(result.pages);
  const pages = result.pages.map(({ id, diagnostics, assessment }, index) => {
    // The leaf publisher does not infer visual support from file existence.
    // Legacy/direct callers without assessment remain explicitly unassessed.
    assessment ??= previewAssessment({ path: `$.pages[${index}]`,
      pageId: typeof id === "string" && id ? id : undefined, diagnostics });
    return { id, assessment, status: assessment.status, diagnostics: assessment.diagnostics, reliability: assessment.reliability };
  });
  const assessment = result.assessment ?? previewAssessment({ diagnostics: result.diagnostics, children: pages.map((page) => page.assessment) });
  const receipt = {
    schema: SCHEMA,
    renderer: result.renderer,
    implementation: await implementationIdentity(result.renderer),
    environment: { node: process.version, raster: { status: "not-loaded" } },
    ...evidence,
    canvas: result.canvas,
    renderEvidence: "local-svg-preview",
    visualReview: "requires-human",
    assessment,
    status: assessment.status,
    reliability: assessment.reliability,
    ok: false,
    diagnostics: assessment.diagnostics,
    output: { directory: destination, status: "incomplete" },
    artifacts: [],
    failures: [],
    pages,
  };
  const fail = (stage, error, pageId, kind) => {
    receipt.failures.push({
      stage, ...(pageId === undefined ? {} : { pageId }), ...(kind ? { kind } : {}),
      message: error instanceof Error ? error.message : String(error),
      ...(error?.code ? { code: String(error.code) } : {}),
    });
    const failedAssessment = pageId === undefined ? receipt.assessment
      : receipt.pages.find((page) => page.id === pageId)?.assessment;
    const diagnostic = previewDiagnostic({
      pageId: typeof pageId === "string" && pageId ? pageId : undefined,
      path: failedAssessment?.path || "$",
      ...(failedAssessment?.scenePath === undefined ? {} : { scenePath: failedAssessment.scenePath }),
      status: "unavailable", reason: stage === "raster-load" ? "raster-dependency-unavailable" : `preview.output.${stage}${kind ? `.${kind}` : ""}`,
      value: error instanceof Error ? error.message : String(error),
      action: "Inspect the publication failure and retained artifacts, repair its cause, and rerun into a new directory.",
    });
    for (const page of receipt.pages) {
      if (pageId !== undefined && page.id !== pageId) continue;
      const next = previewAssessment({ ...page.assessment, assessed: true, diagnostics: [...page.diagnostics, diagnostic] });
      Object.assign(page, { assessment: next, status: next.status, diagnostics: next.diagnostics, reliability: next.reliability });
    }
    // Generated pages can share a semantic root. Preserve both native address
    // and semantic identity instead of replacing every "$" child with one page.
    const identity = (node) => JSON.stringify([node.pageId ?? null, node.id ?? null, node.path, node.scenePath ?? null]);
    const replacements = new Map(receipt.pages.map((page) => [identity(page.assessment), page.assessment]));
    const next = previewAssessment({ ...receipt.assessment, assessed: true,
      diagnostics: [...receipt.diagnostics, diagnostic],
      children: receipt.assessment.children.map((child) => replacements.get(identity(child)) || child) });
    Object.assign(receipt, { assessment: next, status: next.status, diagnostics: next.diagnostics, reliability: next.reliability });
  };
  const outputError = (code, cause) => new PreviewOutputError(
    `PPJ preview output ${code}: ${destination}${cause ? ` (${cause.message || cause})` : ""}`,
    code, receipt, cause,
  );

  try {
    await mkdir(path.dirname(destination), { recursive: true });
    await mkdir(destination);
  } catch (error) {
    fail("destination", error);
    throw outputError(error.code === "EEXIST" ? "preview.output.exists" : "preview.output.reserve", error);
  }
  const pending = path.join(destination, "render.pending.json");
  try {
    await writeArtifact(pending, jsonBytes(receipt));
  } catch (error) {
    fail("pending-manifest", error);
    throw outputError("preview.output.pending", error);
  }

  for (const page of result.pages) {
    for (const diagnostic of page.diagnostics) {
      if (diagnostic.status !== "unavailable") continue;
      fail(diagnostic.reason === "asset-missing" ? "asset" : "preview", diagnostic.reason, page.id);
      if (diagnostic.id !== undefined) receipt.failures.at(-1).elementId = diagnostic.id;
    }
  }
  let raster;
  try {
    raster = await loadRaster();
    if (typeof raster?.render !== "function") throw new Error("Invalid preview raster backend");
    receipt.environment.raster = raster.identity ?? { status: "available", name: "unidentified" };
  } catch (error) {
    raster = null;
    receipt.environment.raster = { name: "sharp", status: "unavailable", reason: error.message || String(error) };
    fail("raster-load", error);
  }

  async function publish(bytes, file, pageIndex, kind) {
    const page = receipt.pages[pageIndex];
    try {
      await writeArtifact(path.join(destination, file), bytes);
      receipt.artifacts.push({ pageId: page.id, kind, file, sha256: sha256(bytes), bytes: bytes.byteLength });
      page[kind === "svg" ? "file" : "png"] = file;
    } catch (error) {
      fail("artifact-write", error, page.id, kind);
    }
  }
  for (let index = 0; index < result.pages.length; index += 1) {
    const page = result.pages[index];
    await publish(Buffer.from(page.svg), `${stems[index]}.svg`, index, "svg");
    if (!raster) continue;
    let bytes;
    try {
      bytes = await raster.render(Buffer.from(page.svg));
      if (!(bytes instanceof Uint8Array) || bytes.byteLength === 0) throw new Error("Raster backend returned no PNG bytes");
    } catch (error) {
      fail("raster-render", error, page.id, "png");
      continue;
    }
    await publish(bytes, `${stems[index]}.png`, index, "png");
  }
  receipt.ok = receipt.failures.length === 0;
  receipt.output.status = receipt.ok ? "complete" : "incomplete";
  try {
    await writeArtifact(path.join(destination, "render.json"), jsonBytes(receipt));
  } catch (error) {
    receipt.ok = false;
    receipt.output.status = "incomplete";
    fail("final-manifest", error);
    throw outputError("preview.output.manifest", error);
  }
  // Final evidence wins over a leftover marker. Do not mutate the persisted
  // receipt on cleanup failure: return a separate warning to the caller.
  const warnings = [];
  try {
    await removePending(pending);
  } catch (error) {
    warnings.push({ stage: "pending-cleanup", message: error.message || String(error), path: pending });
  }
  if (!receipt.ok) {
    const error = outputError("preview.output.incomplete");
    error.warnings = warnings;
    throw error;
  }
  return { receipt, warnings };
}
