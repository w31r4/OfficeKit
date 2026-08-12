#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

export const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
export const COMPACT_SKILL_JSON_PATHS = Object.freeze([
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library/artifact-tool-compose/content-tokens.json",
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library/artifact-tool-compose/template-registry.json",
]);

const MAX_SOURCE_BYTES = 2 * 1024 * 1024;

export function compactJsonText(source, label = "JSON asset") {
  if (typeof source !== "string") throw new TypeError(`${label} source must be a string.`);
  if (Buffer.byteLength(source) > MAX_SOURCE_BYTES) {
    throw new RangeError(`${label} exceeds the ${MAX_SOURCE_BYTES}-byte source budget.`);
  }
  let value;
  try {
    value = JSON.parse(source);
  } catch (error) {
    throw new SyntaxError(`${label} is not valid JSON: ${error.message}`);
  }
  const compact = `${JSON.stringify(value)}\n`;
  if (JSON.stringify(JSON.parse(compact)) !== compact.slice(0, -1)) {
    throw new Error(`${label} did not survive deterministic JSON compaction.`);
  }
  return compact;
}

export async function compactSkillJsons({ check = true, repoRoot = REPO_ROOT } = {}) {
  const changed = [];
  let totalBytes = 0;
  for (const relativePath of COMPACT_SKILL_JSON_PATHS) {
    const filename = path.join(repoRoot, relativePath);
    const stat = await fs.lstat(filename);
    if (!stat.isFile()) throw new Error(`Skill JSON asset must be a regular file: ${relativePath}`);
    const source = await fs.readFile(filename, "utf8");
    const compact = compactJsonText(source, relativePath);
    totalBytes += Buffer.byteLength(compact);
    if (source === compact) continue;
    changed.push({ filename, relativePath, compact, savings: Buffer.byteLength(source) - Buffer.byteLength(compact), mode: stat.mode });
  }

  if (check && changed.length) {
    throw new Error([
      `${changed.length} derived Skill JSON asset(s) are not compact (${changed.reduce((sum, item) => sum + item.savings, 0)} bytes available):`,
      ...changed.map((item) => `- ${item.relativePath}: ${item.savings} bytes`),
      "Run npm run assets:json:compact and review the text-only diff.",
    ].join("\n"));
  }

  if (!check) {
    let serial = 0;
    for (const item of changed) {
      const temporary = `${item.filename}.tmp-${process.pid}-${serial++}`;
      try {
        await fs.writeFile(temporary, item.compact, { mode: item.mode });
        await fs.rename(temporary, item.filename);
      } catch (error) {
        await fs.rm(temporary, { force: true });
        throw error;
      }
    }
  }

  return Object.freeze({ files: COMPACT_SKILL_JSON_PATHS.length, changed: changed.length, totalBytes, savings: changed.reduce((sum, item) => sum + item.savings, 0) });
}

const entry = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (entry === fileURLToPath(import.meta.url)) {
  const flag = process.argv[2] || "--check";
  if (process.argv.length > 3 || !["--check", "--write"].includes(flag)) {
    console.error("Usage: node scripts/compact-skill-jsons.mjs [--check|--write]");
    process.exitCode = 2;
  } else {
    compactSkillJsons({ check: flag !== "--write" }).then((result) => {
      console.log(`Derived Skill JSON compact: ${result.files} files, ${result.totalBytes} bytes${result.changed ? `, ${result.savings} bytes ${flag === "--write" ? "saved" : "available"}` : ""}`);
    }).catch((error) => {
      console.error(error.message);
      process.exitCode = 1;
    });
  }
}
