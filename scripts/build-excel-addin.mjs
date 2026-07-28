#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { cp, mkdir, rm } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

const root = path.resolve(import.meta.dirname, "..");
const app = path.join(root, "apps", "excel-addin");
const output = path.join(app, "dist");
const tsc = path.join(root, "node_modules", "typescript", "bin", "tsc");
const skillIcon = path.join(
  root,
  "skills",
  "spreadsheets",
  "skills",
  "excel-live-control",
  "assets",
  "file-spreadsheet.png",
);

await rm(output, { recursive: true, force: true });
await mkdir(path.join(output, "assets"), { recursive: true });
const result = spawnSync(process.execPath, [tsc, "-p", path.join(app, "tsconfig.json")], {
  cwd: root,
  encoding: "utf8",
});
if (result.status !== 0) {
  process.stderr.write(result.stderr || result.stdout || "TypeScript compilation failed.\n");
  process.exit(result.status ?? 1);
}
for (const filename of ["taskpane.html", "taskpane.css", "support.html"]) {
  await cp(path.join(app, "src", filename), path.join(output, filename));
}
await cp(skillIcon, path.join(output, "assets", "officekit-excel-32.png"));
await cp(skillIcon, path.join(output, "assets", "officekit-excel-80.png"));
