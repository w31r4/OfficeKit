import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const legacyName = /\b(?:openchestnut|openxmlwasm)\b/iu;
const skippedDirectories = new Set([".git", "handoff", "node_modules", "reference", "tmp"]);
const textExtensions = new Set([
  ".c", ".cs", ".csproj", ".css", ".h", ".html", ".js", ".json", ".mjs", ".md", ".proto",
  ".ps1", ".sh", ".txt", ".xml", ".yaml", ".yml",
]);

async function collectTextFiles(directory) {
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    if (entry.isSymbolicLink()) continue;
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      if (!skippedDirectories.has(entry.name)) files.push(...await collectTextFiles(absolute));
      continue;
    }
    if (entry.name === "officekit-identity.mjs") continue;
    if (textExtensions.has(path.extname(entry.name).toLowerCase()) || ["LICENSE", "NOTICE"].includes(entry.name)) files.push(absolute);
  }
  return files;
}

const files = await collectTextFiles(repoRoot);
const matches = [];
for (const filename of files) {
  const text = await fs.readFile(filename, "utf8");
  const match = legacyName.exec(text);
  if (match) {
    const line = text.slice(0, match.index).split("\n").length;
    matches.push(`${path.relative(repoRoot, filename)}:${line}`);
  }
}

assert.deepEqual(matches, [], `public OfficeKit sources retain retired codec names: ${matches.join(", ")}`);

const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
assert.equal(packageJson.name, "office-kit");
assert.equal(packageJson.exports["./codec"], "./src/codecs/office-kit.mjs");
assert.equal(packageJson.exports["./codec/wire"], "./src/generated/office_kit/artifact/v1/office_artifact_pb.js");
assert.ok(await fs.stat(path.join(repoRoot, "native", "OfficeKit")));
assert.ok(await fs.stat(path.join(repoRoot, "runtime", "office-kit")));
assert.ok(await fs.stat(path.join(repoRoot, "proto", "office_kit", "artifact", "v1")));

console.log(`OfficeKit identity gate passed (${files.length} text files scanned)`);
