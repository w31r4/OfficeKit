import { mkdir, readFile, unlink } from "node:fs/promises";
import path from "node:path";
import { sha256, writeExclusiveFile } from "./workspace.mjs";

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

async function implementationIdentity() {
  const files = ["svg-preview.mjs", "preview-output.mjs", "svg-preview-capabilities.json"];
  return {
    version: 1,
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
  const receipt = {
    schema: SCHEMA,
    renderer: result.renderer,
    implementation: await implementationIdentity(),
    environment: { node: process.version, raster: { status: "not-loaded" } },
    ...evidence,
    canvas: result.canvas,
    renderEvidence: "local-svg-preview",
    visualReview: "requires-human",
    status: result.status,
    ok: false,
    diagnostics: [...result.diagnostics],
    output: { directory: destination, status: "incomplete" },
    artifacts: [],
    failures: [],
    pages: result.pages.map(({ id, diagnostics }) => ({ id, diagnostics: [...diagnostics] })),
  };
  const fail = (stage, error, pageId, kind) => {
    receipt.failures.push({
      stage, ...(pageId === undefined ? {} : { pageId }), ...(kind ? { kind } : {}),
      message: error instanceof Error ? error.message : String(error),
      ...(error?.code ? { code: String(error.code) } : {}),
    });
    receipt.status = "unavailable";
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
      fail("asset", diagnostic.reason, page.id);
      receipt.failures.at(-1).elementId = diagnostic.id;
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
    receipt.diagnostics.push({ status: "unavailable", reason: "raster-dependency-unavailable" });
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
