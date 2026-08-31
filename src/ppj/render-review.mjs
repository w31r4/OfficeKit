import { mkdir, rm } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { FileBlob } from "../shared/file-blob.mjs";
import {
  compilePpjWorkspace,
  loadPpjWorkspace,
  sha256,
  writeExclusiveFile,
} from "./workspace.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export async function renderPpj(
  { inputPath, outputPath, pages },
  {
    cwd = process.cwd(),
    load = loadPpjWorkspace,
    compile = compilePpjWorkspace,
    officeRenderer,
    rasterRenderer,
  } = {},
) {
  const workspace = await load(inputPath, { cwd, retainRoot: false });
  const destination = path.resolve(cwd, outputPath);
  if (destination === workspace.path || destination === workspace.sourcePath) {
    throw new Error("PPJ render output must be a directory distinct from the PPJ and source PPTX.");
  }
  await mkdir(path.dirname(destination), { recursive: true });
  try {
    await mkdir(destination);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(`PPJ render output already exists: ${destination}`);
    throw error;
  }

  try {
    const compiled = await compile(workspace, { includeNodeMap: false });
    const program = JSON.parse(Buffer.from(compiled.programJson).toString("utf8"));
    const pageCount = Array.isArray(program.pages) ? program.pages.length : 0;
    const selected = parsePageSelection(pages, pageCount);
    const office = officeRenderer ?? (await import("../renderers/libreoffice.mjs")).createLibreOfficeRenderer({ timeoutMs: 120_000 });
    const raster = rasterRenderer ?? (await import("../renderers/poppler.mjs")).createPopplerRenderer({ dpi: 120, timeoutMs: 120_000 });
    const candidate = new FileBlob(compiled.file, { type: PPTX_MIME, name: `${path.basename(workspace.path, ".ppj")}.pptx` });
    const pdf = await office({
      input: candidate,
      inputType: PPTX_MIME,
      outputType: "application/pdf",
      format: "pdf",
      artifactKind: "presentation",
    });
    const rendered = [];
    for (const page of selected) {
      const png = await raster({
        input: pdf,
        inputType: "application/pdf",
        outputType: "image/png",
        format: "png",
        artifactKind: "presentation",
        pageIndex: page - 1,
      });
      const fileName = `slide-${String(page).padStart(3, "0")}.png`;
      const filePath = path.join(destination, fileName);
      await writeExclusiveFile(filePath, png.bytes, 0o644);
      rendered.push(Object.freeze({ page, file: filePath, sha256: sha256(png.bytes), bytes: png.bytes.byteLength }));
    }
    const evidence = {
      schema: "office-kit/ppj-render/v1",
      input: workspace.path,
      output: destination,
      programSha256: compiled.programSha256,
      candidateSha256: compiled.outputSha256,
      sourceBound: compiled.sourceBound,
      pageCount,
      selectedPages: selected,
      renderer: "libreoffice-poppler",
      renderEvidence: "native-file-render",
      visualReview: "requires-human",
      pages: rendered,
    };
    await writeExclusiveFile(path.join(destination, "render.json"), Buffer.from(`${JSON.stringify(evidence, null, 2)}\n`), 0o644);
    return Object.freeze({ ok: true, command: "render", ...evidence });
  } catch (error) {
    await rm(destination, { recursive: true, force: true });
    throw error;
  }
}

export async function reviewPpj(
  { inputPath, taskId },
  {
    cwd = process.cwd(),
    load = loadPpjWorkspace,
    compile = compilePpjWorkspace,
    review,
  } = {},
) {
  const workspace = await load(inputPath, { cwd, retainRoot: false });
  const compiled = await compile(workspace, { includeNodeMap: true });
  const reviewArtifact = review ?? (await import("../review/index.mjs")).reviewArtifact;
  const report = await reviewArtifact(compiled.file, {
    format: "pptx",
    ppjReceipt: compiled,
    source: workspace.sourcePath ?? undefined,
    playbackEvidence: "structural",
    visualReview: "unavailable",
    contentView: "none",
  });
  const task = taskId == null ? null : await (await import("./task.mjs")).recordPpjTask({
    taskId,
    cwd,
    stage: "reviewed",
    workspace,
    receipt: compiled,
    candidate: { bytes: compiled.file },
    review: report,
  });
  return Object.freeze({
    ok: report.verdict !== "failed",
    command: "review",
    input: workspace.path,
    programSha256: compiled.programSha256,
    candidateSha256: compiled.outputSha256,
    sourceBound: compiled.sourceBound,
    playbackEvidence: "structural",
    visualReview: "unavailable",
    report,
    task,
  });
}

export function parsePageSelection(spec, pageCount) {
  if (!Number.isSafeInteger(pageCount) || pageCount < 1) throw new Error("PPJ render requires at least one page.");
  if (spec == null) return Object.freeze(Array.from({ length: pageCount }, (_, index) => index + 1));
  const values = new Set();
  for (const token of String(spec).split(",")) {
    const value = token.trim();
    const range = /^(\d+)-(\d+)$/u.exec(value);
    if (range) {
      const start = Number(range[1]);
      const end = Number(range[2]);
      if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 1 || end > pageCount) {
        throw new Error(`PPJ page selector must stay within 1..${pageCount}: ${value}`);
      }
      if (start > end) throw new Error(`PPJ page range must be ascending: ${value}`);
      for (let page = start; page <= end; page += 1) values.add(page);
    } else if (/^\d+$/u.test(value)) {
      const page = Number(value);
      if (!Number.isSafeInteger(page)) throw new Error(`Invalid PPJ page selector: ${value}`);
      values.add(page);
    }
    else throw new Error(`Invalid PPJ page selector: ${value || "<empty>"}`);
  }
  const pages = [...values].sort((left, right) => left - right);
  if (pages.length === 0 || pages.some((page) => page < 1 || page > pageCount)) {
    throw new Error(`PPJ page selector must stay within 1..${pageCount}.`);
  }
  return Object.freeze(pages);
}
