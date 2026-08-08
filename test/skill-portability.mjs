import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const skillRoots = [
  "office-kit",
  "documents",
  "spreadsheets",
  "presentations",
  "pdf",
  "template-creator",
].map((name) => path.join(repoRoot, "skills", name));

const ignoredDirectories = new Set([".codex-plugin", "agents", "__pycache__", "reference"]);
const ignoredBasenames = new Set(["LICENSE", "LICENSE.txt", "LICENSE.md"]);
const textExtensions = new Set([".md", ".mjs", ".py", ".txt", ".json", ".yaml", ".yml"]);

async function collectFiles(directory) {
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    if (entry.isDirectory()) {
      if (!ignoredDirectories.has(entry.name)) {
        files.push(...await collectFiles(path.join(directory, entry.name)));
      }
      continue;
    }
    if (ignoredBasenames.has(entry.name) || !textExtensions.has(path.extname(entry.name))) continue;
    files.push(path.join(directory, entry.name));
  }
  return files;
}

const files = (await Promise.all(skillRoots.map(collectFiles))).flat().sort();
assert.ok(files.length > 0, "expected packaged Skill files");

const forbidden = [
  [/mcp__/iu, "MCP tool identifiers"],
  [/:codex-/iu, "host citation directives"],
  [/\$SCRATCH_ROOT/iu, "host scratch variables"],
  [/\$THREAD_ID/iu, "thread variables"],
  [/CODEX_THREAD_ID/iu, "thread variables"],
  [/\bprojectless\b/iu, "host project modes"],
  [/\bproject-backed\b/iu, "host project modes"],
  [/\bload_workspace_dependencies\b/iu, "host dependency loaders"],
  [/\bimage_gen\b/iu, "named image tools"],
  [/google-drive@/iu, "plugin installation instructions"],
  [/\bCodex\b/iu, "host product names"],
  [/ctx\.import\(\s*["'][a-z][a-z0-9+.-]*:\/\//iu, "remote REPL imports"],
];

for (const file of files) {
  const text = await fs.readFile(file, "utf8");
  for (const [pattern, label] of forbidden) {
    assert.doesNotMatch(text, pattern, `${label}: ${path.relative(repoRoot, file)}`);
  }
}

const installRoot = await fs.mkdtemp(path.join(os.tmpdir(), "officekit-skill-portability-"));
try {
  const installed = spawnSync(process.execPath, [
    path.join(repoRoot, "bin", "officekit.mjs"),
    "init",
    installRoot,
    "--tools",
    "agents",
    "--yes",
    "--json",
  ], { cwd: repoRoot, encoding: "utf8" });
  assert.equal(installed.status, 0, installed.stderr);
  const installResult = JSON.parse(installed.stdout);
  assert.deepEqual(installResult.skills, [
    "office-kit",
    "documents",
    "spreadsheets",
    "excel-live-control",
    "presentations",
    "powerpoint-live-control",
    "pdf",
    "template-creator",
  ]);
  for (const skillId of installResult.skills) {
    const skillPath = path.join(installRoot, ".agents", "skills", skillId, "SKILL.md");
    const text = await fs.readFile(skillPath, "utf8");
    for (const [pattern, label] of forbidden) {
      assert.doesNotMatch(text, pattern, `installed ${label}: ${skillId}`);
    }
    assert.match(text, /officekit repl|references\/repl\.md/i, `installed REPL guidance: ${skillId}`);
  }
} finally {
  await fs.rm(installRoot, { recursive: true, force: true });
}

const officeKitRoot = path.join(repoRoot, "skills", "office-kit", "skills", "office-kit");
const workspace = await fs.readFile(path.join(officeKitRoot, "references", "workspace.md"), "utf8");
const capabilities = await fs.readFile(path.join(officeKitRoot, "references", "capabilities.md"), "utf8");
const review = await fs.readFile(path.join(officeKitRoot, "references", "review.md"), "utf8");
const repl = await fs.readFile(path.join(officeKitRoot, "references", "repl.md"), "utf8");
const officeKitSkill = await fs.readFile(path.join(officeKitRoot, "SKILL.md"), "utf8");
const presentationSkill = await fs.readFile(path.join(repoRoot, "skills", "presentations", "skills", "presentations", "SKILL.md"), "utf8");

for (const name of ["workspaceRoot", "taskRoot", "inputRoot", "assetRoot", "outputRoot", "evidenceRoot", "sessionId"]) {
  assert.match(workspace, new RegExp(`\\b${name}\\b`), `workspace contract: ${name}`);
}
assert.match(workspace, /absolute path/i);
assert.match(workspace, /sha256/i);
assert.match(workspace, /visualReview/);
assert.match(workspace, /process\.cwd\(\)/);
assert.match(workspace, /os\.tmpdir\(\)/);
assert.match(workspace, /must not overwrite/i);
assert.match(workspace, /traversal/i);
assert.match(capabilities, /image_view/);
assert.match(capabilities, /image_generate/);
for (const status of ["complete", "unavailable", "requires-human"]) {
  assert.match(capabilities, new RegExp(status));
}
assert.match(capabilities, /native Office shapes/i);
assert.match(capabilities, /\| yes \| yes \|[\s\S]*visualReview: "complete"/i);
assert.match(capabilities, /\| yes \| no \|[\s\S]*user\/template assets/i);
assert.match(capabilities, /\| no \| yes \|[\s\S]*low-risk/i);
assert.match(capabilities, /\| no \| no \|[\s\S]*ask for an asset/i);
assert.match(officeKitSkill, /references\/workspace\.md/);
assert.match(officeKitSkill, /references\/capabilities\.md/);
assert.match(officeKitSkill, /references\/review\.md/);
assert.match(officeKitSkill, /absolute path.*SHA-256/is);
for (const [number, label] of [[6, "Semantic"], [7, "Structural"], [8, "Layout"], [9, "Optional content"], [10, "Visual"], [11, "Delivery"]]) {
  assert.match(review, new RegExp(`${number}\\.\\s+\\*{0,2}${label}`, "i"), `review contract: ${number}. ${label}`);
}
assert.match(review, /contentView: "anydoc"/);
assert.match(review, /do not run AnyDoc merely because\s+it is installed/i);
assert.match(review, /not OCR.*not a substitute for render review/is);
assert.match(review, /visualReview: "requires-human"/);
assert.match(repl, /ctx\.state/);
assert.match(repl, /ctx\.publish/);
assert.match(repl, /maybeApplied/);
assert.match(repl, /not.*replay/is);
assert.match(repl, /process-local/);
assert.match(presentationSkill, /image_view/);
assert.match(presentationSkill, /image_generate/);
assert.match(presentationSkill, /native PowerPoint shapes/i);
assert.match(presentationSkill, /visualReview: "unavailable"/);

for (const [relative, expected] of [
  ["skills/spreadsheets/skills/spreadsheets/routing/google_sheets.md", [/local `\.xlsx`/i, /separate host\s+step/i]],
  ["skills/presentations/skills/presentations/routing/google_slides.md", [/local `\.pptx`/i, /separate host\s+step/i]],
]) {
  const text = await fs.readFile(path.join(repoRoot, relative), "utf8");
  for (const pattern of expected) assert.match(text, pattern, relative);
  assert.doesNotMatch(text, /plugin|MCP|google-drive@|mcp__/i, `${relative} must stay local-file only`);
}

for (const [name, relative] of [
  ["documents", ["documents", "skills", "documents", "SKILL.md"]],
  ["spreadsheets", ["spreadsheets", "skills", "spreadsheets", "SKILL.md"]],
  ["excel-live-control", ["spreadsheets", "skills", "excel-live-control", "SKILL.md"]],
  ["presentations", ["presentations", "skills", "presentations", "SKILL.md"]],
  ["pdf", ["pdf", "skills", "pdf", "SKILL.md"]],
  ["template-creator", ["template-creator", "skills", "template-creator", "SKILL.md"]],
]) {
  const skillPath = path.join(repoRoot, "skills", ...relative);
  const text = await fs.readFile(skillPath, "utf8");
  assert.match(text, /\.\.\/office-kit\/references\/workspace\.md/, `${name} must use the shared contract`);
  assert.match(text, /officekit repl|\.\.\/office-kit\/references\/repl\.md/i, `${name} must use the portable REPL contract`);
}

for (const [name, relative] of [
  ["documents", ["documents", "skills", "documents", "SKILL.md"]],
  ["spreadsheets", ["spreadsheets", "skills", "spreadsheets", "SKILL.md"]],
  ["presentations", ["presentations", "skills", "presentations", "SKILL.md"]],
  ["pdf", ["pdf", "skills", "pdf", "SKILL.md"]],
]) {
  const text = await fs.readFile(path.join(repoRoot, "skills", ...relative), "utf8");
  assert.match(text, /\.\.\/office-kit\/references\/review\.md/, `${name} must use the shared review contract`);
  assert.match(text, /AnyDoc/i, `${name} must describe the optional content view`);
}

console.log(`Skill portability ok: ${files.length} host-neutral files checked`);
