#!/usr/bin/env node
/**
 * Assemble the narrow Windows native closure for the managed Poppler QA pack.
 *
 * This is release tooling, not an installer.  It accepts only a verified,
 * already-extracted upstream source tree and copies three CLI entrypoints,
 * their same-directory DLL closure, and Poppler's data files into a fresh
 * payload.  The generic capability-pack builder then validates the resulting
 * tree again before producing the customer archive.
 */

import { execFile as execFileCallback } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execFile = promisify(execFileCallback);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_INPUTS = path.join(__dirname, "pdf-provider-poppler-release-inputs.v1.json");
const INPUT_SCHEMA = "office-kit.pdf-provider-poppler-release-inputs.v1";
const SHA256 = /^[a-f0-9]{64}$/i;
const PE_MAGIC = Buffer.from([0x4d, 0x5a]);
const PE_SIGNATURE = Buffer.from([0x50, 0x45, 0x00, 0x00]);
const REQUIRED_COMMANDS = ["pdfinfo.exe", "pdftoppm.exe", "pdftotext.exe"];

function fail(message) {
  throw new Error(`Poppler native payload build: ${message}`);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function isSafeSegment(value) {
  return nonEmptyString(value) && !value.includes("/") && !value.includes("\\") && value !== "." && value !== "..";
}

function isSafeRelativePath(value) {
  if (!nonEmptyString(value) || value.includes("\\") || value.startsWith("/")) return false;
  const normalized = path.posix.normalize(value);
  return normalized !== "." && normalized !== ".." && !normalized.startsWith("../") && !normalized.includes("/../");
}

function assertLockedAsset(value, label) {
  if (!value || typeof value !== "object" || !nonEmptyString(value.url) || !SHA256.test(value.sha256 || "")
    || !Number.isSafeInteger(value.downloadBytes) || value.downloadBytes <= 0) {
    fail(`${label} must pin an HTTPS URL, SHA-256, and positive download size.`);
  }
  let url;
  try {
    url = new URL(value.url);
  } catch {
    fail(`${label}.url must be HTTPS.`);
  }
  if (url.protocol !== "https:") fail(`${label}.url must be HTTPS.`);
}

async function loadInputs() {
  const bytes = await fs.readFile(DEFAULT_INPUTS);
  let inputs;
  try {
    inputs = JSON.parse(bytes.toString("utf8"));
  } catch {
    fail("release input lock is not JSON.");
  }
  if (inputs.schema !== INPUT_SCHEMA || inputs.schemaVersion !== 1) fail("release input lock has an unsupported schema.");
  const pack = inputs.popplerQa;
  if (!pack || pack.packId !== "poppler-qa" || pack.version !== "24.08.0-oat.1" || pack.license !== "GPL-2.0-or-later") {
    fail("release input lock must identify the immutable Poppler QA pack.");
  }
  const native = pack.nativeBuild?.["win32-x64"];
  if (!native || native.version !== "24.08.0-0" || !isSafeSegment(native.root)
    || !isSafeRelativePath(native.binRelativePath) || !isSafeRelativePath(native.dataRelativePath)) {
    fail("release input lock has an invalid Windows Poppler source layout.");
  }
  assertLockedAsset(native, "Windows Poppler source");
  assertLockedAsset(inputs.licenseMaterial?.popplerGpl20, "Poppler GPL-2.0 license material");
  return inputs;
}

function parseArguments(argv) {
  if (argv.length === 1 && argv[0] === "--verify-lock") return { verifyLock: true };
  const values = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail(`unexpected argument ${token}.`);
    const name = token.slice(2);
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) fail(`--${name} requires a value.`);
    if (Object.hasOwn(values, name)) fail(`--${name} may be supplied only once.`);
    values[name] = value;
    index += 1;
  }
  for (const required of ["platform", "payload", "bin-root", "data-root", "notices"]) {
    if (!nonEmptyString(values[required])) fail(`--${required} is required.`);
  }
  if (values.platform !== "win32-x64") fail("--platform must be win32-x64.");
  return {
    verifyLock: false,
    platform: values.platform,
    payload: path.resolve(values.payload),
    binRoot: path.resolve(values["bin-root"]),
    dataRoot: path.resolve(values["data-root"]),
    notices: path.resolve(values.notices),
  };
}

async function realDirectory(value, label) {
  const stat = await fs.lstat(value).catch(() => undefined);
  if (!stat?.isDirectory() || stat.isSymbolicLink()) fail(`${label} must be a real directory.`);
  return fs.realpath(value);
}

async function isPeFile(target) {
  const stat = await fs.lstat(target).catch(() => undefined);
  if (!stat?.isFile() || stat.isSymbolicLink() || stat.size < 64) return false;
  const handle = await fs.open(target, "r");
  try {
    const header = Buffer.alloc(64);
    const { bytesRead } = await handle.read(header, 0, header.length, 0);
    if (bytesRead !== header.length || !header.subarray(0, PE_MAGIC.length).equals(PE_MAGIC)) return false;
    const offset = header.readUInt32LE(0x3c);
    if (offset < header.length || offset > 1024 * 1024 || offset + PE_SIGNATURE.length > stat.size) return false;
    const signature = Buffer.alloc(PE_SIGNATURE.length);
    const read = await handle.read(signature, 0, signature.length, offset);
    return read.bytesRead === signature.length && signature.equals(PE_SIGNATURE);
  } finally {
    await handle.close();
  }
}

async function safePeFile(value, label) {
  const stat = await fs.lstat(value).catch(() => undefined);
  if (!stat?.isFile() || stat.isSymbolicLink() || !await isPeFile(value)) fail(`${label} must be a regular PE file.`);
  return fs.realpath(value);
}

async function copyFile(source, destination, { executable = false } = {}) {
  const stat = await fs.lstat(source);
  if (!stat.isFile() || stat.isSymbolicLink()) fail(`source is not a regular file: ${source}.`);
  await fs.mkdir(path.dirname(destination), { recursive: true, mode: 0o755 });
  const existing = await fs.lstat(destination).catch(() => undefined);
  if (existing) fail(`destination already exists: ${destination}.`);
  await fs.copyFile(source, destination);
  await fs.chmod(destination, executable ? 0o755 : 0o644);
}

async function copyDataTree(source, destination) {
  const root = await realDirectory(source, "Poppler data root");
  async function copy(current, target) {
    const stat = await fs.lstat(current);
    if (stat.isSymbolicLink()) fail(`Poppler data contains a symlink: ${current}.`);
    if (stat.isDirectory()) {
      await fs.mkdir(target, { recursive: true, mode: 0o755 });
      const entries = await fs.readdir(current, { withFileTypes: true });
      entries.sort((left, right) => left.name.localeCompare(right.name, "en"));
      for (const entry of entries) {
        if (!isSafeSegment(entry.name)) fail(`Poppler data contains an unsafe entry: ${entry.name}.`);
        await copy(path.join(current, entry.name), path.join(target, entry.name));
      }
      return;
    }
    if (!stat.isFile()) fail(`Poppler data contains an unsupported filesystem entry: ${current}.`);
    await copyFile(current, target);
  }
  await copy(root, destination);
}

async function copyWindowsClosure(binRoot, payloadBin) {
  const root = await realDirectory(binRoot, "Poppler binary root");
  const entries = await fs.readdir(root, { withFileTypes: true });
  entries.sort((left, right) => left.name.localeCompare(right.name, "en"));
  const byName = new Map(entries.map((entry) => [entry.name.toLowerCase(), entry]));
  for (const command of REQUIRED_COMMANDS) {
    const entry = byName.get(command);
    if (!entry || !entry.isFile()) fail(`Poppler binary root is missing ${command}.`);
    const source = await safePeFile(path.join(root, entry.name), command);
    await copyFile(source, path.join(payloadBin, command), { executable: true });
  }
  let dlls = 0;
  for (const entry of entries) {
    if (entry.isDirectory() || !entry.name.toLowerCase().endsWith(".dll")) continue;
    if (!entry.isFile()) fail(`Poppler binary root has an unsupported DLL entry: ${entry.name}.`);
    const source = await safePeFile(path.join(root, entry.name), entry.name);
    await copyFile(source, path.join(payloadBin, entry.name), { executable: true });
    dlls += 1;
  }
  if (!dlls) fail("Poppler binary root has no DLL runtime closure.");
  return dlls;
}

function windowsRuntimeEnvironment(payload) {
  const systemRoot = String(process.env.SystemRoot || process.env.WINDIR || "").trim();
  if (!systemRoot) fail("Windows runtime probe requires SystemRoot.");
  const bin = path.join(payload, "bin");
  return {
    SystemRoot: systemRoot,
    WINDIR: systemRoot,
    PATH: `${bin};${path.join(systemRoot, "System32")}`,
  };
}

async function probeCommand(payload, name) {
  const executable = path.join(payload, "bin", name);
  const { stdout, stderr } = await execFile(executable, ["-v"], {
    env: windowsRuntimeEnvironment(payload),
    timeout: 20_000,
    maxBuffer: 64 * 1024,
    windowsHide: true,
  });
  const output = `${stdout}${stderr}`.trim();
  if (!/poppler/i.test(output)) fail(`${name} did not report a Poppler version.`);
  return output.slice(0, 4096);
}

async function build(options, inputs) {
  if (process.platform !== "win32") fail("win32 payload must be assembled on win32.");
  const payload = await realDirectory(options.payload, "payload");
  const payloadEntries = await fs.readdir(payload);
  if (payloadEntries.length) fail("payload must be an empty directory.");
  const native = inputs.popplerQa.nativeBuild[options.platform];
  const payloadBin = path.join(payload, "bin");
  await fs.mkdir(payloadBin, { recursive: true, mode: 0o755 });
  const dlls = await copyWindowsClosure(options.binRoot, payloadBin);
  await copyDataTree(options.dataRoot, path.join(payload, "share", "poppler"));
  const versions = {};
  for (const command of REQUIRED_COMMANDS) versions[command] = await probeCommand(payload, command);
  if (Object.values(versions).some((output) => !output.includes(native.version.replace(/-0$/, "")))) {
    fail(`runtime version does not match locked Poppler ${native.version}.`);
  }
  const notice = [
    "# Poppler native capability-pack build evidence",
    "",
    `The Windows payload contains Poppler ${native.version} pdfinfo, pdftoppm, and pdftotext plus the exact same-directory DLL closure from the hash-pinned upstream ZIP.`,
    "",
    "## Runtime probes",
    ...Object.entries(versions).flatMap(([name, output]) => [`### ${name}`, "", output, ""]),
    "## Native DLL count",
    "",
    String(dlls),
    "",
  ].join("\n");
  await fs.writeFile(options.notices, notice, { mode: 0o600 });
  process.stdout.write(`${JSON.stringify({ platform: options.platform, payload, dlls, versions }, null, 2)}\n`);
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const inputs = await loadInputs();
  if (options.verifyLock) {
    process.stdout.write(`${JSON.stringify({ schema: inputs.schema, pack: inputs.popplerQa.packId, version: inputs.popplerQa.version, platform: "win32-x64" })}\n`);
    return;
  }
  await build(options, inputs);
}

main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
