import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

const root = path.resolve(import.meta.dirname, "..");
const inputsPath = path.join(root, "scripts", "pdf-provider-verapdf-release-inputs.v1.json");
const builder = path.join(root, "scripts", "build-verapdf-provider-pack.mjs");
const windowsLauncher = path.join(root, "scripts", "windows-java-jar-launcher.c");
const workflowPath = path.join(root, ".github", "workflows", "pdf-verapdf-capability-pack.yml");
const [inputBytes, builderSource, workflowSource] = await Promise.all([
  fs.readFile(inputsPath),
  fs.readFile(builder, "utf8"),
  fs.readFile(workflowPath, "utf8"),
]);
const windowsLauncherSource = await fs.readFile(windowsLauncher, "utf8");
const inputs = JSON.parse(inputBytes);

function run(arguments_, { expect = 0 } = {}) {
  const result = spawnSync(process.execPath, [builder, ...arguments_], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, expect, result.stderr || result.stdout);
  return result;
}

assert.equal(inputs.schema, "office-kit.pdf-provider-verapdf-release-inputs.v1");
assert.equal(inputs.schemaVersion, 1);
assert.equal(inputs.verapdf.packId, "verapdf");
assert.equal(inputs.verapdf.version, "1.30.2-oat.2");
assert.match(inputs.verapdf.license, /MPL-2\.0/);
assert.match(inputs.verapdf.license, /Classpath-exception-2\.0/);
assert.equal(inputs.verapdf.installer.downloadBytes, 32923960);
assert.equal(inputs.verapdf.installer.installerJarPath, "verapdf-greenfield-1.30.2/verapdf-izpack-installer-1.30.2.jar");
assert.equal(inputs.verapdf.installer.cliJarPath, "bin/cli-1.30.2.jar");
assert.equal(inputs.verapdf.installer.archiveFormat, "zip");
assert.deepEqual(Object.keys(inputs.verapdf.jre.platforms).sort(), ["darwin-arm64", "linux-x64", "win32-x64"]);
assert.equal(inputs.verapdf.jre.version, "21.0.7+6");
assert.equal(inputs.verapdf.jre.javaVersion, "21.0.7");
assert.equal(inputs.verapdf.jre.platforms["darwin-arm64"].javaHomePath, "jdk-21.0.7+6-jre/Contents/Home");
assert.equal(inputs.verapdf.jre.platforms["linux-x64"].javaHomePath, "jdk-21.0.7+6-jre");
assert.equal(inputs.verapdf.jre.platforms["win32-x64"].archiveFormat, "zip");
assert.equal(inputs.verapdf.jre.platforms["win32-x64"].javaHomePath, "jdk-21.0.7+6-jre");
for (const source of [
  inputs.verapdf.installer,
  ...Object.values(inputs.verapdf.jre.platforms),
  ...inputs.verapdf.licenseMaterial,
]) {
  assert.match(source.url, /^https:\/\//);
  assert.match(source.sha256, /^[a-f0-9]{64}$/);
  assert.ok(Number.isSafeInteger(source.downloadBytes) && source.downloadBytes > 0);
}
assert.deepEqual(inputs.verapdf.licenseMaterial.map((source) => source.name).sort(), [
  "veraPDF GPL-3.0-or-later",
  "veraPDF MPL-2.0",
]);

const summary = JSON.parse(run(["--verify-lock"]).stdout);
assert.equal(summary.sha256, crypto.createHash("sha256").update(inputBytes).digest("hex"));
assert.equal(summary.pack, "verapdf");
assert.equal(summary.platforms["darwin-arm64"].downloadBytes, 80951034);
assert.equal(summary.platforms["linux-x64"].downloadBytes, 84787557);
assert.equal(summary.version, "1.30.2-oat.2");
assert.equal(summary.platforms["win32-x64"].downloadBytes, 81799320);

for (const sourceFragment of [
  "downloadPinned",
  "copyDereferencedTree",
  "safeJavaHome",
  "safeJavaEnvironment",
  "installer template",
  "jre/legal/",
  "source contains a dangling symlink",
  "source contains a symlink directory cycle",
  "env -i HOME=",
  "JAVA_HOME",
  "veraPDF CLI JAR",
  "pack launcher does not expose the required built-in ua2 profile",
  "GUI, samples, and uninstaller are not packaged",
  "--windows-launcher is required for win32-x64",
  "archiveFormat must be tar.gz or zip",
  "Windows native launcher",
  "cli.jar",
]) assert.match(builderSource, new RegExp(sourceFragment.replace(/[.*+?^\x24{}()|[\]\\]/g, "\\$&")));
assert.doesNotMatch(builderSource, /HOME: process\.env\.HOME/);
assert.match(workflowSource, /verify-pdf-provider-pack\.mjs/);
assert.doesNotMatch(workflowSource, /launcher_args|tar -xzf/);
assert.match(workflowSource, /OFFICE_KIT_PDF_VERAPDF_TEST =/,
  "the Windows adapter smoke must receive the native launcher path through PowerShell");
assert.match(workflowSource, /node\.exe test\/pdf-verapdf-provider\.mjs/,
  "the Windows adapter smoke must launch Node without Git Bash path conversion");

for (const sourceFragment of [
  "CreateProcessW",
  "JAVA_TOOL_OPTIONS",
  "JDK_JAVA_OPTIONS",
  "_JAVA_OPTIONS",
  "CLASSPATH",
  "jre\\\\bin\\\\java.exe",
  "libexec\\\\bin\\\\cli.jar",
  "CommandLineToArgvW",
]) assert.match(windowsLauncherSource, new RegExp(sourceFragment.replace(/[.*+?^\x24{}()|[\]\\]/g, "\\$&")));
assert.doesNotMatch(windowsLauncherSource, /cmd\.exe|powershell|system\s*\(/i, "the pack launcher must not delegate to a command interpreter");

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-verapdf-lock-"));
try {
  const unsafe = structuredClone(inputs);
  unsafe.verapdf.jre.platforms["linux-x64"].javaHomePath = "../escape";
  const unsafePath = path.join(temporary, "unsafe.json");
  await fs.writeFile(unsafePath, JSON.stringify(unsafe), "utf8");
  const rejected = run(["--verify-lock", "--inputs", unsafePath], { expect: 2 });
  assert.match(rejected.stderr, /safe relative path/);
  const invalidArchive = structuredClone(inputs);
  invalidArchive.verapdf.jre.platforms["win32-x64"].archiveFormat = "exe";
  const invalidArchivePath = path.join(temporary, "invalid-archive.json");
  await fs.writeFile(invalidArchivePath, JSON.stringify(invalidArchive), "utf8");
  assert.match(run(["--verify-lock", "--inputs", invalidArchivePath], { expect: 2 }).stderr, /archiveFormat must be tar\.gz or zip/);
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("veraPDF PDF capability-pack build smoke ok");
