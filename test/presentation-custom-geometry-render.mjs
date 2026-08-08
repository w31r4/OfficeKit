import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { Presentation, PresentationFile } from "../src/index.mjs";

const PROBE_TIMEOUT_MS = 5_000;
const RENDER_TIMEOUT_MS = 120_000;

function commandAvailable(command, args = ["--version"]) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    timeout: PROBE_TIMEOUT_MS,
    killSignal: "SIGKILL",
    maxBuffer: 64 * 1024,
  });
  return !result.error && result.status === 0;
}

function run(command, args) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    timeout: RENDER_TIMEOUT_MS,
    killSignal: "SIGKILL",
    maxBuffer: 1024 * 1024,
  });
  assert.equal(result.error, undefined, `${command} could not start: ${result.error?.message || "unknown error"}`);
  assert.equal(result.status, 0, `${command} failed with ${result.status}: ${String(result.stderr || result.stdout).trim()}`);
  return result;
}

function probePositions(bboxHtml) {
  return [...bboxHtml.matchAll(/<page\b[^>]*>([\s\S]*?)<\/page>/g)].map((page, index) => {
    const word = /<word\b([^>]*)>RECTPROBE<\/word>/.exec(page[1]);
    assert.ok(word, `LibreOffice/Poppler page ${index + 1} must expose RECTPROBE as selectable text.`);
    const x = Number(/\bxMin="([^"]+)"/.exec(word[1])?.[1]);
    const y = Number(/\byMin="([^"]+)"/.exec(word[1])?.[1]);
    assert.ok(Number.isFinite(x) && Number.isFinite(y), `Page ${index + 1} needs finite RECTPROBE coordinates.`);
    return { x, y };
  });
}

if (!commandAvailable("soffice") || !commandAvailable("pdftotext", ["-v"])) {
  console.log("presentation custom geometry native render skipped (LibreOffice/Poppler unavailable)");
  process.exit(0);
}

const work = await fs.mkdtemp(path.join(os.tmpdir(), "officekit-custom-text-rectangle-"));
try {
  const presentation = Presentation.create({ slideSize: { width: 640, height: 360 } });
  for (const [name, textRectangle] of [["Default text bounds", undefined], ["Inset text bounds", { left: 180, top: 90, right: 520, bottom: 240 }]]) {
    const slide = presentation.slides.add({ name });
    slide.shapes.add({
      name: `probe-${slide.index + 1}`,
      geometry: "custom",
      position: { left: 40, top: 40, width: 560, height: 280 },
      fill: "#DBEAFE",
      line: { fill: "#2563EB", width: 2 },
      text: "RECTPROBE",
      textRectangle,
      textStyle: { fontFamily: "Liberation Sans", fontSize: 24, color: "#0F172A" },
      textBodyProperties: { anchor: "top", insets: { left: 0, top: 0, right: 0, bottom: 0 } },
      customPaths: [{
        width: 100,
        height: 100,
        commands: [
          { moveTo: { x: 50, y: 0 } },
          { lineTo: { x: 100, y: 50 } },
          { lineTo: { x: 50, y: 100 } },
          { lineTo: { x: 0, y: 50 } },
          { close: {} },
        ],
      }],
    });
  }

  const pptxPath = path.join(work, "custom-text-rectangle.pptx");
  await (await PresentationFile.exportPptx(presentation)).save(pptxPath);
  const profilePath = path.join(work, "libreoffice-profile");
  run("soffice", [
    `-env:UserInstallation=${pathToFileURL(profilePath).href}`,
    "--headless",
    "--convert-to", "pdf:impress_pdf_Export",
    "--outdir", work,
    pptxPath,
  ]);
  const pdfPath = path.join(work, "custom-text-rectangle.pdf");
  const bboxPath = path.join(work, "custom-text-rectangle.html");
  assert.ok((await fs.stat(pdfPath)).size > 0, "LibreOffice must produce a non-empty PDF.");
  run("pdftotext", ["-bbox-layout", pdfPath, bboxPath]);
  const positions = probePositions(await fs.readFile(bboxPath, "utf8"));
  assert.equal(positions.length, 2, "Native PDF must retain both presentation slides.");
  assert.ok(positions[1].x - positions[0].x > 80, `Inset text must move right in native output: ${JSON.stringify(positions)}`);
  assert.ok(positions[1].y - positions[0].y > 35, `Inset text must move down in native output: ${JSON.stringify(positions)}`);
  console.log(`presentation custom geometry native render ok ${JSON.stringify(positions)}`);
} finally {
  await fs.rm(work, { recursive: true, force: true });
}
