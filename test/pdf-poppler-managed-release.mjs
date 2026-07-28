import assert from "node:assert/strict";
import { execFile as execFileCallback } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { PdfProviders } from "../src/pdf/providers/index.mjs";
import { plainPdfBytes } from "./fixtures/plain-pdf.mjs";

const execFile = promisify(execFileCallback);

function combinedOutput(result) {
  return String(result.stdout || "") + String(result.stderr || "");
}

if (process.env.OFFICE_KIT_PDF_LIVE_PACK_TEST !== "1") {
  console.log("Poppler managed release smoke skipped (set OFFICE_KIT_PDF_LIVE_PACK_TEST=1)");
} else if (`${process.platform}-${process.arch}` !== "win32-x64") {
  console.log("Poppler managed release smoke skipped (the published Poppler QA pack is win32-x64 only)");
} else {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-poppler-managed-release-"));
  try {
    const policyDirectory = path.join(temporary, ".office-kit");
    const policyPath = path.join(policyDirectory, "pdf-providers.json");
    await fs.mkdir(policyDirectory);
    await fs.writeFile(policyPath, JSON.stringify({
      installPolicy: "managed",
      allowedProviders: ["poppler"],
      allowedPacks: ["poppler-qa"],
      acceptedLicenses: [],
      allowedOcrLanguages: ["eng", "chi_sim"],
      maxDownloadBytes: 32 * 1024 * 1024,
      maxUnpackedBytes: 64 * 1024 * 1024,
    }), "utf8");

    const source = path.join(temporary, "source.pdf");
    const sourceBytes = Buffer.from(plainPdfBytes([{ text: "managed Poppler QA smoke", width: 612, height: 792 }]));
    await fs.writeFile(source, sourceBytes, { mode: 0o600 });
    const resolution = await PdfProviders.resolve({
      task: "render",
      provider: "poppler",
      inspection: { summary: { sourceSha256: "c".repeat(64) } },
      savePolicy: "read-only",
      policyPath,
    });
    assert.equal(resolution.status, "installable", JSON.stringify(resolution.reason));
    assert.deepEqual(resolution.installPlan?.packIds, ["poppler-qa"]);
    assert.equal(resolution.installPlan?.packs?.[0]?.artifact?.platform, "win32-x64");

    const ready = await PdfProviders.ensure({ resolution, policyPath });
    assert.equal(ready.status, "ready", JSON.stringify(ready.reason));
    const runtime = ready.runtime?.managed;
    const pdfinfo = runtime?.commandPaths?.pdfinfo;
    const pdftoppm = runtime?.commandPaths?.pdftoppm;
    const pdftotext = runtime?.commandPaths?.pdftotext;
    assert.ok(pdfinfo && pdftoppm && pdftotext, "managed Poppler must return every catalogued command path");
    for (const executable of [pdfinfo, pdftoppm, pdftotext]) {
      assert.match(executable, /\\bin\\[^\\]+\.exe$/i, "managed Poppler must expose Windows archive paths, never ambient commands");
      const stat = await fs.lstat(executable);
      assert.ok(stat.isFile() && !stat.isSymbolicLink(), "managed Poppler executable must be a regular private-cache file");
    }

    const commandOptions = { timeout: 15_000, maxBuffer: 64 * 1024, windowsHide: true, env: { ...process.env, ...runtime.environment } };
    assert.match(combinedOutput(await execFile(pdfinfo, ["-v"], commandOptions)), /poppler/i);
    assert.match(combinedOutput(await execFile(pdfinfo, [source], commandOptions)), /Pages:\s+1\b/i);
    const extracted = await execFile(pdftotext, [source, "-"], commandOptions);
    assert.match(combinedOutput(extracted), /managed Poppler QA smoke/);
    const rasterPrefix = path.join(temporary, "rendered");
    await execFile(pdftoppm, ["-singlefile", "-png", "-r", "72", source, rasterPrefix], commandOptions);
    assert.ok((await fs.stat(`${rasterPrefix}.png`)).size > 0, "managed pdftoppm must produce a native raster");
    assert.deepEqual(await fs.readFile(source), sourceBytes, "read-only Poppler QA must not mutate its source");

    const cached = await PdfProviders.probe({ provider: "poppler", task: "render", policyPath });
    assert.equal(cached.status, "ready", JSON.stringify(cached.reason));
    assert.equal(cached.runtime?.managed?.commandPaths?.pdfinfo, pdfinfo);
    assert.equal(cached.runtime?.managed?.commandPaths?.pdftoppm, pdftoppm);
    assert.equal(cached.runtime?.managed?.commandPaths?.pdftotext, pdftotext);
    console.log("Poppler managed release smoke ok (" + ready.installation.installed["poppler-qa"].receipt.artifact.asset + ")");
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}
