import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";

const root = path.resolve(import.meta.dirname, "..");
const inputsPath = path.join(root, "scripts", "pdf-provider-ocr-release-inputs.v1.json");
const pythonInputsPath = path.join(root, "scripts", "pdf-provider-python-release-inputs.v1.json");
const nativeBuilder = path.join(root, "scripts", "build-ocr-native-payload.mjs");
const workflowPath = path.join(root, ".github", "workflows", "pdf-ocr-capability-packs.yml");
const windowsPythonLauncher = path.join(root, "scripts", "windows-python-module-launcher.c");
const windowsGhostscriptLauncher = path.join(root, "scripts", "windows-ghostscript-launcher.c");
const windowsSidecarLauncher = path.join(root, "scripts", "windows-ocr-sidecar-launcher.c");

const [inputBytes, pythonInputBytes, nativeSource, workflowSource, windowsPythonLauncherSource, windowsGhostscriptLauncherSource, windowsSidecarLauncherSource] = await Promise.all([
  fs.readFile(inputsPath),
  fs.readFile(pythonInputsPath),
  fs.readFile(nativeBuilder, "utf8"),
  fs.readFile(workflowPath, "utf8"),
  fs.readFile(windowsPythonLauncher, "utf8"),
  fs.readFile(windowsGhostscriptLauncher, "utf8"),
  fs.readFile(windowsSidecarLauncher, "utf8"),
]);
const inputs = JSON.parse(inputBytes);

assert.equal(inputs.schema, "office-kit.pdf-provider-ocr-release-inputs.v1");
assert.equal(inputs.schemaVersion, 1);
assert.equal(inputs.ocrCore.packId, "ocr-core");
assert.equal(inputs.ocrCore.version, "17.8.1-oat.3");
assert.equal(
  inputs.ocrCore.pythonInputs.sha256,
  crypto.createHash("sha256").update(pythonInputBytes).digest("hex"),
  "the OCR release lock must pin the exact isolated-Python wheel lock it builds",
);
assert.deepEqual(inputs.ocrCore.nativeBuild["darwin-arm64"].formulae, ["tesseract", "ghostscript", "poppler"]);
assert.deepEqual(inputs.ocrCore.nativeBuild["linux-x64"].packages, ["tesseract-ocr", "tesseract-ocr-eng", "ghostscript", "poppler-utils", "poppler-data", "libgs-common", "fonts-droid-fallback", "fonts-urw-base35", "patchelf"]);
const windowsNative = inputs.ocrCore.nativeBuild["win32-x64"];
assert.equal(windowsNative.tesseract.version, "5.5.0.20241111");
assert.equal(windowsNative.ghostscript.version, "10.05.1");
assert.equal(windowsNative.poppler.version, "24.08.0-0");
assert.equal(windowsNative.poppler.root, "poppler-24.08.0");
assert.match(windowsNative.tesseract.url, /tesseract-ocr-w64-setup-5\.5\.0\.20241111\.exe$/);
assert.match(windowsNative.ghostscript.url, /gs10051w64\.exe$/);
assert.match(windowsNative.poppler.url, /Release-24\.08\.0-0\.zip$/);
for (const source of Object.values(windowsNative)) {
  assert.match(source.url, /^https:\/\//);
  assert.match(source.sha256, /^[a-f0-9]{64}$/);
  assert.ok(Number.isSafeInteger(source.downloadBytes) && source.downloadBytes > 0);
}

for (const [language, expected] of Object.entries({ eng: "ocr-language-eng", chi_sim: "ocr-language-chi-sim" })) {
  const languageInput = inputs.languages[language];
  assert.equal(languageInput.packId, expected);
  assert.equal(languageInput.version, "4.1.0-oat.3");
  assert.equal(languageInput.license, "Apache-2.0");
  assert.match(languageInput.url, /^https:\/\//);
  assert.match(languageInput.sha256, /^[a-f0-9]{64}$/);
  assert.ok(Number.isSafeInteger(languageInput.downloadBytes) && languageInput.downloadBytes > 0);
}
assert.match(inputs.licenseMaterial.tessdataFastApache20.url, /^https:\/\//);
assert.match(inputs.licenseMaterial.tessdataFastApache20.sha256, /^[a-f0-9]{64}$/);
assert.ok(inputs.licenseMaterial.tessdataFastApache20.downloadBytes > 1000);
for (const source of Object.values(inputs.licenseMaterial.windowsNative)) {
  assert.match(source.url, /^https:\/\//);
  assert.match(source.sha256, /^[a-f0-9]{64}$/);
  assert.ok(Number.isSafeInteger(source.downloadBytes) && source.downloadBytes > 1000);
}
for (const licenseName of ["ghostscriptAgpl30", "popplerGpl20"]) {
  assert.match(
    inputs.licenseMaterial.windowsNative[licenseName].url,
    /^https:\/\/raw\.githubusercontent\.com\/spdx\/license-list-data\/[a-f0-9]{40}\/text\//,
    `${licenseName} must use an immutable, hash-pinned SPDX license-text source rather than an availability-sensitive live license endpoint`,
  );
}

// The release builder must force every package-local executable through its
// own relocated libraries, remove all bundled language data, and leave only a
// separately authorized language-pack directory for the provider adapter.
for (const sourceFragment of [
  "removeTraineddata",
  "DYLD_FALLBACK_LIBRARY_PATH",
  "LD_LIBRARY_PATH",
  "GS_LIB",
  "codesign",
  "patchelf",
  "writeLaunchers",
  "native library basename collision",
  "MACHO_MAGICS",
  "isMachOFile",
  "ELF_MAGIC",
  "isElfFile",
  "PE_MAGIC",
  "isPeFile",
  "copyWindowsSidecars",
  "copyWindowsSidecarLibraries",
  "private closure has conflicting",
  "windows-sidecar-root",
  "Windows library root contains a non-PE DLL",
  "windows-python-launcher",
  "windows-ghostscript-launcher",
  "windows-tesseract-launcher",
  "windows-pdftotext-launcher",
  "loaderPath",
  "linuxRpath",
  "listLinuxLibraryFiles",
  "LD_LIBRARY_PATH: libDirectory",
  "listMacLibraryFiles",
  "path.basename(source)",
  "native library destination is unsafe",
  "contains a dangling symlink",
  "contains a symlink directory cycle",
  "resource-root",
]) assert.match(nativeSource, new RegExp(sourceFragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.match(nativeSource, /if \(!await isMachOFile\(target\)\) return false;/);
assert.match(nativeSource, /path\.relative\(path\.dirname\(target\), destination\)/);
assert.match(workflowSource, /fonts-droid-fallback/);
assert.match(workflowSource, /platform: linux-x64\s+# Jammy ships Tesseract 4\.1[\s\S]*?runner: ubuntu-24\.04/);
assert.match(workflowSource, /--resource-root/);
assert.match(workflowSource, /fonts-urw-base35/);
assert.match(workflowSource, /libgs-common/);
assert.match(workflowSource, /poppler-data/);
assert.match(workflowSource, /platform: win32-x64\s+runner: windows-2025/);
assert.match(workflowSource, /windows-python-module-launcher\.c/);
assert.match(workflowSource, /windows-ghostscript-launcher\.c/);
assert.match(workflowSource, /windows-ocr-sidecar-launcher\.c/);
assert.match(workflowSource, /office-kit-build-ocr-launchers\.cmd/);
assert.match(workflowSource, /VsDevCmd\.bat/);
assert.match(workflowSource, /OFFICEKIT_TESSERACT_LAUNCHER/);
assert.match(workflowSource, /OFFICEKIT_PDFTOTEXT_LAUNCHER/);
assert.match(workflowSource, /--windows-sidecar-root/);
assert.match(workflowSource, /native\/tesseract\/tesseract\.exe/);
assert.match(workflowSource, /native\/ghostscript\/gswin64c\.exe/);
assert.match(workflowSource, /native\/poppler\/pdftotext\.exe/);
assert.match(workflowSource, /tesseract v\?5\\\./,
  "the cross-platform runtime probe must accept the official Windows Tesseract v5 banner");
assert.match(workflowSource, /Expand-Archive/);
assert.match(workflowSource, /Extract-PinnedTesseractRuntime/,
  "the Windows Tesseract source must be extracted into the private build root instead of installed globally");
assert.match(workflowSource, /Get-Command 7z\.exe/,
  "the Windows Tesseract extractor must use the hosted 7-Zip tool only against the hash-pinned source archive");
assert.match(workflowSource, /never run it or let it choose a machine-wide install path/,
  "the Windows Tesseract path must not execute the installer or use a machine-wide destination");
assert.match(workflowSource, /Extract-PinnedGhostscriptRuntime/,
  "the Windows Ghostscript source must be extracted into the private build root instead of installed globally");
assert.match(workflowSource, /Ghostscript archive exposed gswin64c\.exe without executing the installer/,
  "the Windows Ghostscript path must prove it did not execute the self-extracting installer");
assert.match(workflowSource, /7-Zip could not extract the pinned Ghostscript archive/,
  "a malformed or unsupported Ghostscript source must fail with an exact private-extraction diagnostic");
assert.match(workflowSource, /The extracted Ghostscript archive did not expose gswin64c\.exe/,
  "the private Ghostscript extraction must require the exact native CLI payload");
assert.doesNotMatch(workflowSource, /Start-Process|WaitForExit|Stop-Process|Start-Sleep/,
  "the Windows Ghostscript source must never execute an installer or poll a child process");
const windowsPowerShellBlock = workflowSource.match(/powershell\.exe -NoProfile -NonInteractive -Command '\n([\s\S]*?)\n            '/)?.[1];
assert.ok(windowsPowerShellBlock, "the Windows OCR release path must retain one explicit PowerShell block");
assert.doesNotMatch(windowsPowerShellBlock, /'/,
  "the single-quoted Bash PowerShell payload must not contain an unescaped apostrophe that breaks every platform's shell parse");
assert.match(workflowSource, /--expected-platforms darwin-arm64,linux-x64,win32-x64/);
assert.match(workflowSource, /\.catalogFragment\.artifacts \| length == 3/);
assert.match(workflowSource, /resource_target/);
assert.match(workflowSource, /resource_actual/);
assert.match(workflowSource, /dpkg-query -S/);
assert.match(workflowSource, /for root_formula in tesseract ghostscript poppler; do brew deps/);
assert.match(workflowSource, /unapproved Ghostscript resource target/);
assert.match(workflowSource, /run_probe\(\)/);
assert.match(workflowSource, /ocr-home/);
assert.match(workflowSource, /sha256-file\.mjs/);
assert.match(workflowSource, /verify-pdf-provider-pack\.mjs/);
assert.doesNotMatch(workflowSource, /brew deps --include-optional/);
assert.doesNotMatch(workflowSource, /brew deps --union tesseract ghostscript poppler/);
assert.doesNotMatch(workflowSource, /shasum/);

for (const source of [windowsPythonLauncherSource, windowsGhostscriptLauncherSource]) {
  for (const fragment of ["CreateProcessW", "CommandLineToArgvW", "SystemRoot", "PATH"]) {
    assert.match(source, new RegExp(fragment));
  }
  assert.doesNotMatch(source, /cmd\.exe|powershell|system\s*\(/i, "the Windows OCR launchers must not delegate to a command interpreter");
}
assert.match(windowsPythonLauncherSource, /parent_directory\(root\) \|\| !parent_directory\(root\)/);
assert.match(windowsGhostscriptLauncherSource, /GS_LIB/);
assert.match(windowsGhostscriptLauncherSource, /native\\\\ghostscript/);
assert.match(windowsGhostscriptLauncherSource, /gswin64c\.exe/);
for (const sourceFragment of [
  "CreateProcessW",
  "CommandLineToArgvW",
  "OFFICEKIT_TESSERACT_LAUNCHER",
  "OFFICEKIT_PDFTOTEXT_LAUNCHER",
  "native\\\\tesseract",
  "native\\\\poppler",
  "TESSDATA_PREFIX",
]) assert.match(windowsSidecarLauncherSource, new RegExp(sourceFragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.doesNotMatch(windowsSidecarLauncherSource, /cmd\.exe|powershell|system\s*\(/i, "the private sidecar launcher must not delegate to a command interpreter");

const invalidPlatform = spawnSync(process.execPath, [nativeBuilder,
  "--platform", "win32-arm64",
  "--payload", root,
  "--notices", path.join(root, "package.json"),
  "--tesseract", process.execPath,
  "--ghostscript", process.execPath,
  "--pdftotext", process.execPath,
  "--ghostscript-root", root,
  "--tessdata-root", root,
  "--library-root", root,
], { cwd: root, encoding: "utf8" });
assert.equal(invalidPlatform.status, 2);
assert.match(invalidPlatform.stderr, /platform must be one of darwin-arm64, linux-x64, win32-x64/);

const missingWindowsLaunchers = spawnSync(process.execPath, [nativeBuilder,
  "--platform", "win32-x64",
  "--payload", root,
  "--notices", path.join(root, "package.json"),
  "--tesseract", process.execPath,
  "--ghostscript", process.execPath,
  "--pdftotext", process.execPath,
  "--ghostscript-root", root,
  "--tessdata-root", root,
  "--library-root", root,
], { cwd: root, encoding: "utf8" });
assert.equal(missingWindowsLaunchers.status, 2);
assert.match(missingWindowsLaunchers.stderr, /--windows-python-launcher is required for win32-x64/);

const missingWindowsSidecars = spawnSync(process.execPath, [nativeBuilder,
  "--platform", "win32-x64",
  "--payload", root,
  "--notices", path.join(root, "package.json"),
  "--tesseract", process.execPath,
  "--ghostscript", process.execPath,
  "--pdftotext", process.execPath,
  "--ghostscript-root", root,
  "--tessdata-root", root,
  "--windows-python-launcher", root,
  "--windows-ghostscript-launcher", root,
  "--windows-tesseract-launcher", root,
  "--windows-pdftotext-launcher", root,
], { cwd: root, encoding: "utf8" });
assert.equal(missingWindowsSidecars.status, 2);
assert.match(missingWindowsSidecars.stderr, /--windows-sidecar-root tesseract=<directory> is required for win32-x64/);

console.log("OCR PDF capability-pack build smoke ok");
