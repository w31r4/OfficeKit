import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";

const root = path.resolve(import.meta.dirname, "..");
const inputsPath = path.join(root, "scripts", "pdf-provider-poppler-release-inputs.v1.json");
const builderPath = path.join(root, "scripts", "build-poppler-native-payload.mjs");
const portableBuilderPath = path.join(root, "scripts", "build-poppler-portable-payload.mjs");
const workflowPath = path.join(root, ".github", "workflows", "pdf-poppler-capability-pack.yml");

const [inputBytes, builder, portableBuilder, workflow] = await Promise.all([
  fs.readFile(inputsPath),
  fs.readFile(builderPath, "utf8"),
  fs.readFile(portableBuilderPath, "utf8"),
  fs.readFile(workflowPath, "utf8"),
]);
const inputs = JSON.parse(inputBytes);
const native = inputs.popplerQa.nativeBuild["win32-x64"];

assert.equal(inputs.schema, "office-kit.pdf-provider-poppler-release-inputs.v1");
assert.equal(inputs.schemaVersion, 1);
assert.equal(inputs.popplerQa.packId, "poppler-qa");
assert.equal(inputs.popplerQa.version, "24.08.0-oat.2");
assert.equal(inputs.popplerQa.license, "GPL-2.0-or-later");
assert.deepEqual(inputs.popplerQa.source, {
  version: "24.08.0",
  url: "https://poppler.freedesktop.org/poppler-24.08.0.tar.xz",
  sha256: "97453fbddf0c9a9eafa0ea45ac710d3d49bcf23a62e864585385d3c0b4403174",
  downloadBytes: 1912592,
});
assert.equal(native.version, "24.08.0-0");
assert.equal(native.root, "poppler-24.08.0");
assert.equal(native.binRelativePath, "Library/bin");
assert.equal(native.dataRelativePath, "share/poppler");
assert.equal(native.licenseRelativePath, "share/poppler/COPYING.gpl2");
assert.equal(native.licenseSha256, "ab15fd526bd8dd18a9e77ebc139656bf4d33e97fc7238cd11bf60e2b9b8666c6");
assert.equal(native.licenseBytes, 17987);
assert.match(native.url, /^https:\/\//);
assert.match(native.sha256, /^[a-f0-9]{64}$/);
assert.ok(Number.isSafeInteger(native.downloadBytes) && native.downloadBytes > 1_000);

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
  "safe archive-relative path, SHA-256, and positive byte size",
  "license material must live inside the copied data tree",
  "pin the Poppler source archive",
]) assert.match(builder, new RegExp(fragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));

for (const fragment of [
  "darwin-arm64 payload must be assembled on macOS",
  "linux-x64 payload must be assembled on Linux",
  "pdfinfo",
  "pdftoppm",
  "pdftotext",
  "otool",
  "install_name_tool",
  "codesign",
  "ldd",
  "patchelf",
  "LD_LIBRARY_PATH",
  "DYLD_LIBRARY_PATH",
  "declared native roots",
  "contains a symlink",
  "must not be hard-linked",
  "POPPLER_DATADIR",
]) assert.match(portableBuilder, new RegExp(fragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));

for (const fragment of [
  "name: Poppler QA (win32-x64)",
  "poppler-posix",
  "ubuntu-24.04",
  "macos-14",
  "runs-on: windows-2025",
  "Expand-Archive",
  "build-poppler-native-payload.mjs",
  "build-poppler-portable-payload.mjs",
  "source_dir=\"$RUNNER_TEMP/poppler-source/poppler-$(jq",
  "-DENABLE_LIBTIFF=OFF",
  "--expected-platforms darwin-arm64,linux-x64,win32-x64",
  "actions/attest@v4",
  "pdf-provider-poppler-qa-",
  "System32",
  "sha256-file.mjs",
  "verify-pdf-provider-pack.mjs",
  "COPYING.gpl2",
]) assert.match(workflow, new RegExp(fragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.doesNotMatch(workflow, /(?:choco|winget|scoop)\s+(?:install|add)/i, "the release lane must not use a global Windows package manager");
assert.doesNotMatch(workflow, /shasum|tar -xzf/, "the Windows release lane must not inherit Git Bash hashing or tar behavior");

const lock = spawnSync(process.execPath, [builderPath, "--verify-lock"], { cwd: root, encoding: "utf8" });
assert.equal(lock.status, 0, lock.stderr || lock.stdout);
assert.deepEqual(JSON.parse(lock.stdout), {
  schema: inputs.schema,
  pack: "poppler-qa",
  version: "24.08.0-oat.2",
  platform: "win32-x64",
});

const missing = spawnSync(process.execPath, [builderPath, "--platform", "win32-x64"], { cwd: root, encoding: "utf8" });
assert.equal(missing.status, 2);
assert.match(missing.stderr, /--payload is required/);

const missingPortable = spawnSync(process.execPath, [portableBuilderPath, "--platform", "darwin-arm64"], { cwd: root, encoding: "utf8" });
assert.equal(missingPortable.status, 2);
assert.match(missingPortable.stderr, /--payload is required/);

assert.equal(crypto.createHash("sha256").update(inputBytes).digest("hex").length, 64);
console.log("Poppler PDF capability-pack build smoke ok");
