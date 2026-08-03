#!/usr/bin/env node
/**
 * Assemble a relocatable Poppler QA payload for macOS or Linux.
 *
 * The input binaries and data tree must already be installed by the reviewed
 * release workflow.  This helper never invokes a package manager, follows a
 * symlink, or copies a dependency outside the explicitly declared roots.  It
 * copies the three read-only QA commands, resolves their non-system native
 * closure, patches loader paths, and proves the staged commands work after
 * extraction before the generic PDF pack builder signs the archive.
 */

import { execFile as execFileCallback } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";

const execFile = promisify(execFileCallback);
const PLATFORMS = new Set(["darwin-arm64", "linux-x64"]);
const REQUIRED_COMMANDS = ["pdfinfo", "pdftoppm", "pdftotext"];
const SHA256 = /^[a-f0-9]{64}$/i;
const MACHO_MAGICS = new Set([
  0xfeedface, 0xcefaedfe, 0xfeedfacf, 0xcffaedfe,
  0xcafebabe, 0xbebafeca, 0xcafebabf, 0xbfbafeca,
]);
const ELF_MAGIC = Buffer.from([0x7f, 0x45, 0x4c, 0x46]);

function fail(message) {
  throw new Error(`Poppler portable payload: ${message}`);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function safeRelativePath(value) {
  if (!nonEmptyString(value) || value.includes("\\") || value.startsWith("/")) return false;
  const normalized = path.posix.normalize(value);
  return normalized !== "." && normalized !== ".." && !normalized.startsWith("../") && !normalized.includes("/../");
}

function parseArguments(argv) {
  const values = {};
  const repeated = new Map([["library-root", []], ["resource-root", []]]);
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail(`unexpected argument ${token}`);
    const name = token.slice(2);
    if (name === "summary") {
      if (values.summary) fail("--summary may be supplied only once");
      values.summary = true;
      continue;
    }
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) fail(`--${name} requires a value`);
    index += 1;
    if (repeated.has(name)) repeated.get(name).push(value);
    else if (Object.hasOwn(values, name)) fail(`--${name} may be supplied only once`);
    else values[name] = value;
  }
  for (const required of ["platform", "payload", "pdfinfo", "pdftoppm", "pdftotext", "data-root", "notices"]) {
    if (!nonEmptyString(values[required])) fail(`--${required} is required`);
  }
  if (!PLATFORMS.has(values.platform)) fail("--platform must be darwin-arm64 or linux-x64");
  if (!repeated.get("library-root").length) fail("at least one --library-root is required");
  if (values["source-sha256"] !== undefined && !SHA256.test(values["source-sha256"])) fail("--source-sha256 must be 64 hexadecimal characters");
  return {
    platform: values.platform,
    payload: path.resolve(values.payload),
    commands: Object.fromEntries(REQUIRED_COMMANDS.map((name) => [name, path.resolve(values[name])])),
    dataRoot: path.resolve(values["data-root"]),
    libraryRoots: repeated.get("library-root").map((value) => path.resolve(value)),
    resourceRoots: repeated.get("resource-root").map((value) => path.resolve(value)),
    notices: path.resolve(values.notices),
    version: values.version?.trim() || undefined,
    sourceUrl: values["source-url"]?.trim() || undefined,
    sourceSha256: values["source-sha256"]?.toLowerCase(),
    summary: values.summary === true,
  };
}

async function lstatReal(file, label) {
  const stat = await fs.lstat(file).catch((error) => fail(`${label} is unavailable: ${file}: ${error.message}`));
  if (stat.isSymbolicLink()) fail(`${label} must not be a symlink: ${file}`);
  return stat;
}

async function realFile(file, label, { executable = false } = {}) {
  const actual = await fs.realpath(file).catch((error) => fail(`${label} is unavailable: ${file}: ${error.message}`));
  const stat = await lstatReal(actual, label);
  if (!stat.isFile() || (executable && (stat.mode & 0o111) === 0)) fail(`${label} must be a regular ${executable ? "executable " : ""}file: ${file}`);
  return actual;
}

async function realDirectory(directory, label) {
  const actual = await fs.realpath(directory).catch((error) => fail(`${label} is unavailable: ${directory}: ${error.message}`));
  const stat = await lstatReal(actual, label);
  if (!stat.isDirectory()) fail(`${label} must be a real directory: ${directory}`);
  return actual;
}

function contained(root, target, label) {
  const relative = path.relative(root, target);
  if (path.isAbsolute(relative) || relative === ".." || relative.startsWith(`..${path.sep}`)) fail(`${label} escapes its declared root: ${target}`);
}

async function isNativeFile(file, platform) {
  const stat = await fs.lstat(file).catch(() => undefined);
  if (!stat?.isFile() || stat.isSymbolicLink() || stat.size < 4) return false;
  const handle = await fs.open(file, "r");
  try {
    const header = Buffer.alloc(4);
    const { bytesRead } = await handle.read(header, 0, header.length, 0);
    if (bytesRead !== header.length) return false;
    return platform === "darwin-arm64" ? MACHO_MAGICS.has(header.readUInt32BE(0)) : header.equals(ELF_MAGIC);
  } finally {
    await handle.close();
  }
}

async function copyRegular(source, destination, label, { executable = false } = {}) {
  const stat = await lstatReal(source, label);
  if (!stat.isFile()) fail(`${label} is not a regular file: ${source}`);
  if (stat.nlink > 1) fail(`${label} must not be hard-linked: ${source}`);
  await fs.mkdir(path.dirname(destination), { recursive: true, mode: 0o755 });
  const existing = await fs.lstat(destination).catch(() => undefined);
  if (existing) {
    if (existing.isSymbolicLink() || !existing.isFile()) fail(`destination is unsafe: ${destination}`);
    const [oldBytes, newBytes] = await Promise.all([fs.readFile(destination), fs.readFile(source)]);
    if (!oldBytes.equals(newBytes)) fail(`dependency basename collision has different bytes: ${path.basename(source)}`);
    return false;
  }
  await fs.copyFile(source, destination, fs.constants.COPYFILE_EXCL);
  await fs.chmod(destination, executable ? 0o755 : 0o644);
  return true;
}

async function copyTree(source, destination, label) {
  await realDirectory(source, label);
  async function walk(current, relative = "") {
    const entries = await fs.readdir(current, { withFileTypes: true });
    entries.sort((left, right) => left.name.localeCompare(right.name, "en"));
    for (const entry of entries) {
      const nextRelative = relative ? `${relative}/${entry.name}` : entry.name;
      if (!safeRelativePath(nextRelative)) fail(`${label} contains unsafe path ${nextRelative}`);
      const sourcePath = path.join(current, entry.name);
      const destinationPath = path.join(destination, ...nextRelative.split("/"));
      const stat = await fs.lstat(sourcePath);
      if (stat.isSymbolicLink()) fail(`${label} contains a symlink: ${nextRelative}`);
      if (entry.isDirectory()) await walk(sourcePath, nextRelative);
      else if (entry.isFile()) {
        if (stat.nlink > 1) fail(`${label} contains a hard-linked file: ${nextRelative}`);
        await copyRegular(sourcePath, destinationPath, `${label} file`, { executable: (stat.mode & 0o111) !== 0 });
      } else fail(`${label} contains an unsupported entry: ${nextRelative}`);
    }
  }
  await fs.mkdir(destination, { recursive: true, mode: 0o755 });
  await walk(source);
}

async function listFiles(root, predicate) {
  const result = [];
  async function walk(current) {
    const entries = await fs.readdir(current, { withFileTypes: true });
    entries.sort((left, right) => left.name.localeCompare(right.name, "en"));
    for (const entry of entries) {
      const file = path.join(current, entry.name);
      const stat = await fs.lstat(file);
      if (stat.isSymbolicLink()) fail(`declared native root contains a symlink: ${file}`);
      if (entry.isDirectory()) await walk(file);
      else if (entry.isFile() && predicate(file, stat)) result.push(file);
      else if (!entry.isFile()) fail(`declared native root contains an unsupported entry: ${file}`);
    }
  }
  await walk(root);
  return result;
}

async function listMacLibraryCandidates(root) {
  const result = [];
  async function walk(current) {
    const entries = await fs.readdir(current, { withFileTypes: true });
    entries.sort((left, right) => left.name.localeCompare(right.name, "en"));
    for (const entry of entries) {
      const candidate = path.join(current, entry.name);
      const stat = await fs.lstat(candidate);
      if (entry.isDirectory()) {
        await walk(candidate);
        continue;
      }
      if (!entry.isFile() && !entry.isSymbolicLink()) fail(`macOS native root contains an unsupported entry: ${candidate}`);
      if (!entry.name.includes(".dylib")) continue;
      const actual = await fs.realpath(candidate).catch((error) => fail(`macOS native library link is unresolved: ${candidate}: ${error.message}`));
      contained(root, actual, "macOS native library link");
      const actualStat = await fs.lstat(actual);
      if (!actualStat.isFile() || actualStat.isSymbolicLink()) fail(`macOS native library link does not resolve to a regular file: ${candidate}`);
      if (actualStat.nlink > 1) fail(`macOS native root contains a hard-linked file: ${actual}`);
      result.push({ name: entry.name, file: actual });
    }
  }
  await walk(root);
  return result;
}

async function run(command, args, label, options = {}) {
  try {
    return await execFile(command, args, { encoding: "utf8", maxBuffer: 512 * 1024, ...options });
  } catch (error) {
    if (options.allowFailure) return undefined;
    fail(`${label} failed: ${String(error?.stderr || error?.stdout || error?.message || error).trim()}`);
  }
}

function parseMacDependencies(output) {
  return String(output).split(/\r?\n/).slice(1).map((line) => line.trim().split(" (")[0]).filter(Boolean);
}

function parseLinuxDependencies(output) {
  const result = [];
  for (const line of String(output).split(/\r?\n/)) {
    const mapped = /=>\s+(\/[^\s]+)\s+\(/.exec(line);
    if (mapped) result.push(mapped[1]);
    else {
      const direct = /^\s*(\/[^\s]+)\s+\(/.exec(line);
      if (direct) result.push(direct[1]);
    }
    if (/=>\s+not found/.test(line)) fail(`native Linux dependency is missing: ${line.trim()}`);
  }
  return result;
}

function hostLinuxLibrary(name) {
  return /^(?:ld-linux|ld-musl|libc\.so|libm\.so|libdl\.so|libpthread\.so|librt\.so|libresolv\.so|libnsl\.so)/.test(name);
}

function macLoaderPath(target, destination) {
  const relative = path.relative(path.dirname(target), destination).split(path.sep).join("/");
  return relative && relative !== "." ? `@loader_path/${relative}` : `@loader_path/${path.basename(destination)}`;
}

function linuxRpath(target, libraryDirectory) {
  const relative = path.relative(path.dirname(target), libraryDirectory).split(path.sep).join("/");
  return relative && relative !== "." ? `$ORIGIN/${relative}` : "$ORIGIN";
}

async function findByBasename(name, roots, label) {
  const matches = [];
  for (const root of roots) {
    const files = (await listMacLibraryCandidates(root)).filter((entry) => entry.name === name).map((entry) => entry.file);
    matches.push(...files);
  }
  if (!matches.length) fail(`${label} ${name} was not found in declared native roots`);
  const bytes = await fs.readFile(matches[0]);
  for (const candidate of matches.slice(1)) {
    if (!bytes.equals(await fs.readFile(candidate))) fail(`${label} basename collision has different bytes: ${name}`);
  }
  return matches[0];
}

async function collectMacLibraries(targets, roots, destination) {
  const copied = new Map();
  const queue = [...targets];
  const seen = new Set();
  while (queue.length) {
    const target = queue.shift();
    const real = await fs.realpath(target);
    if (seen.has(real)) continue;
    seen.add(real);
    const listed = await run("otool", ["-L", target], `otool ${target}`);
    for (const dependency of parseMacDependencies(listed.stdout)) {
      if (dependency.startsWith("/System/") || dependency.startsWith("/usr/lib/") || dependency.startsWith("/usr/lib/swift/")) continue;
      const name = path.basename(dependency);
      const source = await findByBasename(name, roots, "macOS dependency");
      const targetPath = path.join(destination, name);
      await copyRegular(source, targetPath, "macOS dependency");
      copied.set(name, targetPath);
      queue.push(targetPath);
    }
  }
  return copied;
}

async function patchMacPayload(payload, libraryNames) {
  const bin = path.join(payload, "bin");
  const lib = path.join(payload, "lib");
  const targets = [...await listFiles(bin, () => true), ...await listFiles(lib, () => true)];
  for (const target of targets) {
    if (!await isNativeFile(target, "darwin-arm64")) continue;
    const listed = await run("otool", ["-L", target], `otool ${target}`);
    for (const dependency of parseMacDependencies(listed.stdout)) {
      const name = path.basename(dependency);
      if (!libraryNames.has(name)) {
        if (dependency.startsWith("/System/") || dependency.startsWith("/usr/lib/")) continue;
        fail(`relocatable macOS payload retains an undeclared dependency ${dependency}`);
      }
      const replacement = macLoaderPath(target, path.join(lib, name));
      if (dependency !== replacement) await run("install_name_tool", ["-change", dependency, replacement, target], `install_name_tool ${target}`);
    }
    if (path.dirname(target) === lib && path.basename(target).includes(".dylib")) {
      await run("install_name_tool", ["-id", `@loader_path/${path.basename(target)}`, target], `install_name_tool id ${target}`);
    }
    await run("codesign", ["--force", "--sign", "-", "--timestamp=none", target], `codesign ${target}`);
  }
  for (const target of targets) {
    if (!await isNativeFile(target, "darwin-arm64")) continue;
    const listed = await run("otool", ["-L", target], `verify otool ${target}`);
    for (const dependency of parseMacDependencies(listed.stdout)) {
      if (dependency.startsWith("/System/") || dependency.startsWith("/usr/lib/")) continue;
      if (dependency.includes("/opt/homebrew/") || dependency.includes("/usr/local/")) fail(`relocatable macOS payload retains build-machine path ${dependency}`);
      if (dependency.startsWith("@rpath/") && libraryNames.has(path.basename(dependency))) fail(`relocatable macOS payload retains unresolved rpath ${dependency}`);
    }
  }
  return targets.length;
}

async function collectLinuxLibraries(targets, destination, roots) {
  const copied = new Map();
  const queue = [...targets];
  const seen = new Set();
  const rootPaths = await Promise.all([destination, ...roots].map((root) => fs.realpath(root)));
  const isDeclared = (file) => rootPaths.some((root) => {
    const relative = path.relative(root, file);
    return !path.isAbsolute(relative) && relative !== ".." && !relative.startsWith(`..${path.sep}`);
  });
  while (queue.length) {
    const target = queue.shift();
    const real = await fs.realpath(target);
    if (seen.has(real)) continue;
    seen.add(real);
    const listed = await run("ldd", [target], `ldd ${target}`, { env: { ...process.env, LD_LIBRARY_PATH: [destination, ...roots].join(path.delimiter) } });
    for (const dependency of parseLinuxDependencies(listed.stdout)) {
      if (hostLinuxLibrary(path.basename(dependency))) continue;
      const name = path.basename(dependency);
      const targetPath = path.join(destination, name);
      const actual = await fs.realpath(dependency).catch((error) => fail(`Linux dependency is unresolved: ${dependency}: ${error.message}`));
      if (!isDeclared(actual)) fail(`Linux dependency escapes declared native roots: ${actual}`);
      await copyRegular(actual, targetPath, "Linux dependency");
      copied.set(name, targetPath);
      queue.push(targetPath);
    }
  }
  if (!copied.size) fail("no non-system Linux dependencies were resolved");
  return copied;
}

async function patchLinuxPayload(payload) {
  const bin = path.join(payload, "bin");
  const lib = path.join(payload, "lib");
  const targets = [...await listFiles(bin, () => true), ...await listFiles(lib, () => true)];
  for (const target of targets) {
    if (await isNativeFile(target, "linux-x64")) await run("patchelf", ["--set-rpath", linuxRpath(target, lib), target], `patchelf ${target}`);
  }
  return targets.length;
}

async function captureVersion(executable, args, environment) {
  const result = await run(executable, args, `${path.basename(executable)} version probe`, { env: { ...process.env, ...environment, LANG: "C", LC_ALL: "C" } });
  return String(result.stdout || result.stderr || "").trim().slice(0, 4096);
}

async function build(options) {
  if (options.platform === "darwin-arm64" && process.platform !== "darwin") fail("darwin-arm64 payload must be assembled on macOS");
  if (options.platform === "linux-x64" && process.platform !== "linux") fail("linux-x64 payload must be assembled on Linux");
  for (const name of REQUIRED_COMMANDS) options.commands[name] = await realFile(options.commands[name], `${name} executable`, { executable: true });
  options.dataRoot = await realDirectory(options.dataRoot, "Poppler data root");
  options.libraryRoots = await Promise.all(options.libraryRoots.map((root) => realDirectory(root, "native library root")));
  const existing = await fs.readdir(options.payload).catch(() => []);
  if (existing.length) fail("--payload must be an empty directory");
  const bin = path.join(options.payload, "bin");
  const lib = path.join(options.payload, "lib");
  const share = path.join(options.payload, "share", "poppler");
  await Promise.all([fs.mkdir(bin, { recursive: true, mode: 0o755 }), fs.mkdir(lib, { recursive: true, mode: 0o755 })]);
  for (const name of REQUIRED_COMMANDS) await copyRegular(options.commands[name], path.join(bin, name), `${name} executable`, { executable: true });
  await copyTree(options.dataRoot, share, "Poppler data root");
  const targets = REQUIRED_COMMANDS.map((name) => path.join(bin, name));
  let libraryCount;
  let targetCount;
  if (options.platform === "darwin-arm64") {
    const libraries = await collectMacLibraries(targets, options.libraryRoots, lib);
    targetCount = await patchMacPayload(options.payload, new Set(libraries.keys()));
    libraryCount = libraries.size;
  } else {
    const libraries = await collectLinuxLibraries(targets, lib, options.libraryRoots);
    targetCount = await patchLinuxPayload(options.payload);
    libraryCount = libraries.size;
  }
  const environment = options.platform === "darwin-arm64"
    ? { DYLD_LIBRARY_PATH: lib, POPPLER_DATADIR: share }
    : { LD_LIBRARY_PATH: lib, POPPLER_DATADIR: share };
  const versions = {};
  for (const name of REQUIRED_COMMANDS) versions[name] = await captureVersion(path.join(bin, name), ["-v"], environment);
  if (options.version && Object.values(versions).some((value) => !value.includes(options.version))) fail(`staged Poppler commands did not report locked version ${options.version}`);
  const notice = await fs.readFile(options.notices).catch((error) => fail(`--notices is unavailable: ${error.message}`));
  if (!notice.length) fail("--notices must not be empty");
  const summary = { platform: options.platform, payload: options.payload, libraryCount, targetCount, versions, source: options.sourceUrl && options.sourceSha256 ? { url: options.sourceUrl, sha256: options.sourceSha256 } : null };
  return summary;
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const summary = await build(options);
  process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
