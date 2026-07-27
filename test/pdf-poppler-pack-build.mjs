import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";

const root = path.resolve(import.meta.dirname, "..");
const inputsPath = path.join(root, "scripts", "pdf-provider-poppler-release-inputs.v1.json");
const builderPath = path.join(root, "scripts", "build-poppler-native-payload.mjs");
const workflowPath = path.join(root, ".github", "workflows", "pdf-poppler-capability-pack.yml");

const [inputBytes, builder, workflow] = await Promise.all([
  fs.readFile(inputsPath),
  fs.readFile(builderPath, "utf8"),
  fs.readFile(workflowPath, "utf8"),
]);
const inputs = JSON.parse(inputBytes);
const native = inputs.popplerQa.nativeBuild["win32-x64"];

assert.equal(inputs.schema, "office-kit.pdf-provider-poppler-release-inputs.v1");
assert.equal(inputs.schemaVersion, 1);
assert.equal(inputs.popplerQa.packId, "poppler-qa");
assert.equal(inputs.popplerQa.version, "24.08.0-oat.1");
assert.equal(inputs.popplerQa.license, "GPL-2.0-or-later");
assert.equal(native.version, "24.08.0-0");
assert.equal(native.root, "poppler-24.08.0");
assert.equal(native.binRelativePath, "Library/bin");
assert.equal(native.dataRelativePath, "share/poppler");
for (const value of [native, inputs.licenseMaterial.popplerGpl20]) {
  assert.match(value.url, /^https:\/\//);
  assert.match(value.sha256, /^[a-f0-9]{64}$/);
  assert.ok(Number.isSafeInteger(value.downloadBytes) && value.downloadBytes > 1_000);
}

for (const fragment of [
  "win32 payload must be assembled on win32",
  "pdfinfo.exe",
  "pdftoppm.exe",
  "pdftotext.exe",
  "PE_SIGNATURE",
  "Windows runtime probe requires SystemRoot",
  "Poppler data contains a symlink",
  "Poppler binary root has no DLL runtime closure",
  "same-directory DLL closure",
]) assert.match(builder, new RegExp(fragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));

for (const fragment of [
  "name: Poppler QA (win32-x64)",
  "runs-on: windows-2025",
  "Expand-Archive",
  "build-poppler-native-payload.mjs",
  "--expected-platforms win32-x64",
  "actions/attest@v4",
  "pdf-provider-poppler-qa-",
  "System32",
  "sha256-file.mjs",
  "verify-pdf-provider-pack.mjs",
]) assert.match(workflow, new RegExp(fragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.doesNotMatch(workflow, /(?:choco|winget|scoop)\s+(?:install|add)/i, "the release lane must not use a global Windows package manager");
assert.doesNotMatch(workflow, /shasum|tar -xzf/, "the Windows release lane must not inherit Git Bash hashing or tar behavior");

const lock = spawnSync(process.execPath, [builderPath, "--verify-lock"], { cwd: root, encoding: "utf8" });
assert.equal(lock.status, 0, lock.stderr || lock.stdout);
assert.deepEqual(JSON.parse(lock.stdout), {
  schema: inputs.schema,
  pack: "poppler-qa",
  version: "24.08.0-oat.1",
  platform: "win32-x64",
});

const missing = spawnSync(process.execPath, [builderPath, "--platform", "win32-x64"], { cwd: root, encoding: "utf8" });
assert.equal(missing.status, 2);
assert.match(missing.stderr, /--payload is required/);

assert.equal(crypto.createHash("sha256").update(inputBytes).digest("hex").length, 64);
console.log("Poppler PDF capability-pack build smoke ok");
