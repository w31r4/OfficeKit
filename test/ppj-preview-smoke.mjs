import { execFileSync } from "node:child_process";
import { mkdtemp } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { renderPpjToSvg } from "../src/ppj/svg-preview.mjs";

const files = execFileSync("rg", ["--files", "-g", "*.ppj", "-g", "!node_modules/**", "-g", "!tmp/**"], { encoding: "utf8" }).trim().split("\n").filter(Boolean);
const out = await mkdtemp(path.join(os.tmpdir(), "officekit-ppj-preview-smoke-"));
const report = { schema: "office-kit/ppj-preview-smoke/v1", total: files.length, rendered: [], compilerRejected: [], rendererFailed: [] };
for (const file of files) {
  try {
    const result = await renderPpjToSvg(file, { outputDir: path.join(out, String(report.rendered.length)) });
    report.rendered.push({ file, status: result.status, diagnostics: result.diagnostics.length });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/unsupported_ppj_|ppj\.(source|schema)\./u.test(message)) report.compilerRejected.push({ file, error: message });
    else report.rendererFailed.push({ file, error: message });
  }
}
console.log(JSON.stringify(report, null, 2));
if (report.rendererFailed.length) process.exitCode = 1;
