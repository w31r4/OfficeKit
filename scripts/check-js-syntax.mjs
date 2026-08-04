#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const roots = ["src", "scripts", "test", "bin"];
const extensions = new Set([".js", ".mjs", ".cjs"]);

function walk(root) {
  const directory = path.join(repoRoot, root);
  if (!fs.existsSync(directory)) return [];
  return fs.readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const relative = path.join(root, entry.name);
    if (entry.isDirectory()) return walk(relative);
    return extensions.has(path.extname(entry.name)) ? [relative] : [];
  });
}

const files = roots.flatMap(walk).sort();
for (const file of files) {
  const result = spawnSync(process.execPath, ["--check", file], {
    cwd: repoRoot,
    encoding: "utf8",
    stdio: "pipe",
  });
  if (result.status !== 0) {
    process.stderr.write(result.stderr || result.stdout || `syntax check failed: ${file}\n`);
    process.exit(result.status || 1);
  }
}
console.log(`JavaScript syntax/import preflight ok: ${files.length} files`);
