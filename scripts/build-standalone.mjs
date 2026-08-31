#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import { existsSync } from "node:fs";
import {
  access,
  chmod,
  cp,
  lstat,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  rename,
  rm,
  writeFile,
} from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { fileURLToPath, pathToFileURL } from "node:url";
import JSZip from "jszip";
import pako from "pako";

const REPOSITORY_ROOT = path.resolve(import.meta.dirname, "..");
const RUNTIME_CATALOG_PATH = path.join(
  REPOSITORY_ROOT,
  "standalone",
  "node-runtimes.v1.json",
);
const SUPPORTED_TARGETS = new Set(["darwin-arm64", "linux-x64", "win32-x64"]);
const BLOCK_SIZE = 512;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;
const MAX_RUNTIME_ARCHIVE_BYTES = 80_000_000;
const ZIP_TIMESTAMP = new Date("1980-01-01T00:00:00.000Z");

function fail(message) {
  throw new Error(`OfficeKit standalone build: ${message}`);
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function isWindowsTarget(target) {
  return target.startsWith("win32-");
}

function nodeRuntimeTarget(target) {
  return target === "win32-x64" ? "win-x64" : target;
}

function standaloneArchiveExtension(target) {
  return isWindowsTarget(target) ? ".zip" : ".tar.gz";
}

function stableJson(value) {
  return `${JSON.stringify(value, null, 2)}\n`;
}

function lexicalCompare(left, right) {
  return left < right ? -1 : left > right ? 1 : 0;
}

function safeRelativePath(value) {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.includes("\\") ||
    value.includes("\0") ||
    value.startsWith("/")
  ) {
    return false;
  }
  const normalized = path.posix.normalize(value);
  return (
    normalized !== "." &&
    normalized !== ".." &&
    !normalized.startsWith("../") &&
    !normalized.includes("/../")
  );
}

function safeSegment(value, label) {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value === "." ||
    value === ".." ||
    value.includes("/") ||
    value.includes("\\") ||
    value.includes("\0")
  ) {
    fail(`${label} must be one safe path segment.`);
  }
  return value;
}

function parseArguments(argv) {
  const values = {};
  const flags = new Set();
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail(`unexpected argument ${token}.`);
    const name = token.slice(2);
    if (name === "force") {
      if (flags.has(name)) fail("--force may be supplied only once.");
      flags.add(name);
      continue;
    }
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) fail(`--${name} requires a value.`);
    if (Object.hasOwn(values, name)) fail(`--${name} may be supplied only once.`);
    values[name] = value;
    index += 1;
  }
  if (!values.target) fail("--target is required.");
  if (!values.output) fail("--output is required.");
  if (!SUPPORTED_TARGETS.has(values.target)) {
    fail(`--target must be one of ${[...SUPPORTED_TARGETS].join(", ")}.`);
  }
  return {
    target: values.target,
    outputDirectory: path.resolve(values.output),
    runtimeCacheDirectory: values["runtime-cache"]
      ? path.resolve(values["runtime-cache"])
      : path.join(os.homedir(), ".cache", "office-kit", "node"),
    force: flags.has("force"),
  };
}

function run(command, args, options = {}) {
  const executable = process.platform === "win32" && command === "npm"
    ? "npm.cmd"
    : command;
  const result = spawnSync(executable, args, {
    cwd: options.cwd ?? REPOSITORY_ROOT,
    encoding: "utf8",
    env: options.env ?? process.env,
    maxBuffer: 64 * 1024 * 1024,
    shell: process.platform === "win32" && executable === "npm.cmd",
  });
  if (result.error) fail(`${executable} could not start: ${result.error.message}`);
  if (result.status !== 0) {
    const detail = (result.stderr || result.stdout || "").trim();
    fail(`${executable} exited ${result.status}${detail ? `: ${detail}` : "."}`);
  }
  return result;
}

async function regularFile(pathname, label) {
  const metadata = await lstat(pathname);
  if (!metadata.isFile() || metadata.isSymbolicLink()) {
    fail(`${label} must be a regular non-symlink file: ${pathname}.`);
  }
  return metadata;
}

async function validateRuntimeEntry(entry, target, nodeVersion) {
  if (!entry || typeof entry !== "object") fail(`runtime ${target} is missing.`);
  safeSegment(entry.archive, `${target} archive`);
  safeSegment(entry.root, `${target} root`);
  if (!SHA256_PATTERN.test(entry.sha256)) fail(`${target} runtime SHA-256 is invalid.`);
  if (!Number.isSafeInteger(entry.size) || entry.size <= 0 || entry.size > MAX_RUNTIME_ARCHIVE_BYTES) {
    fail(`${target} runtime size is outside the accepted bound.`);
  }
  const expectedRoot = `node-v${nodeVersion}-${nodeRuntimeTarget(target)}`;
  if (entry.root !== expectedRoot) {
    fail(`${target} runtime root must be ${expectedRoot}.`);
  }
  let url;
  try {
    url = new URL(entry.url);
  } catch {
    fail(`${target} runtime URL is invalid.`);
  }
  if (url.protocol !== "https:" || url.hostname !== "nodejs.org") {
    fail(`${target} runtime URL must be an official nodejs.org HTTPS URL.`);
  }
}

async function runtimeCatalog() {
  const catalog = JSON.parse(await readFile(RUNTIME_CATALOG_PATH, "utf8"));
  if (catalog.schemaVersion !== 1) fail("unsupported Node runtime catalog schema.");
  safeSegment(catalog.nodeVersion, "Node version");
  for (const target of SUPPORTED_TARGETS) {
    await validateRuntimeEntry(catalog.runtimes?.[target], target, catalog.nodeVersion);
  }
  return catalog;
}

async function verifyPinnedFile(pathname, expected, label) {
  const metadata = await regularFile(pathname, label);
  if (metadata.size !== expected.size) {
    fail(`${label} size is ${metadata.size}; expected ${expected.size}.`);
  }
  const bytes = await readFile(pathname);
  const actualHash = sha256(bytes);
  if (actualHash !== expected.sha256) {
    fail(`${label} SHA-256 is ${actualHash}; expected ${expected.sha256}.`);
  }
  return bytes;
}

async function downloadPinnedRuntime(entry, cacheDirectory) {
  await mkdir(cacheDirectory, { recursive: true });
  const destination = path.join(cacheDirectory, entry.archive);
  try {
    await access(destination);
    await verifyPinnedFile(destination, entry, "cached Node runtime");
    return destination;
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }

  const response = await fetch(entry.url, { redirect: "follow" });
  if (!response.ok) fail(`Node runtime download returned HTTP ${response.status}.`);
  const declaredLength = Number(response.headers.get("content-length"));
  if (Number.isFinite(declaredLength) && declaredLength !== entry.size) {
    fail(`Node runtime Content-Length is ${declaredLength}; expected ${entry.size}.`);
  }
  if (response.body == null) fail("Node runtime download returned an empty body.");
  const chunks = [];
  let received = 0;
  for await (const chunk of response.body) {
    const bytes = Buffer.from(chunk);
    received += bytes.length;
    if (received > entry.size || received > MAX_RUNTIME_ARCHIVE_BYTES) {
      await response.body.cancel().catch(() => {});
      fail("Node runtime exceeded its pinned download size.");
    }
    chunks.push(bytes);
  }
  const bytes = Buffer.concat(chunks, received);
  if (bytes.length !== entry.size || sha256(bytes) !== entry.sha256) {
    fail("downloaded Node runtime failed its pinned size or SHA-256 check.");
  }
  const temporary = `${destination}.tmp-${process.pid}-${crypto.randomUUID()}`;
  try {
    await writeFile(temporary, bytes, { flag: "wx", mode: 0o600 });
    await rename(temporary, destination);
  } finally {
    await rm(temporary, { force: true });
  }
  return destination;
}

function validateTarListing(listing, expectedRoot) {
  const entries = listing.split(/\r?\n/u).filter(Boolean);
  if (entries.length === 0) fail("archive contains no entries.");
  for (const raw of entries) {
    const entry = raw.replace(/\/$/u, "");
    if (!safeRelativePath(entry)) fail(`archive contains unsafe path ${JSON.stringify(raw)}.`);
    if (entry !== expectedRoot && !entry.startsWith(`${expectedRoot}/`)) {
      fail(`archive entry is outside ${expectedRoot}: ${raw}.`);
    }
  }
}

function zipEntryMode(entry) {
  if (typeof entry.unixPermissions === "number") return entry.unixPermissions;
  if (typeof entry.unixPermissions === "string") {
    return Number.parseInt(entry.unixPermissions, 8);
  }
  return null;
}

function safeZipEntryPath(value) {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    value.includes("\\") ||
    value.includes("\0") ||
    value.startsWith("/") ||
    /^[a-z]:/iu.test(value)
  ) {
    return false;
  }
  return safeRelativePath(value);
}

async function extractPinnedZip(archive, expectedRoot, destination) {
  const zip = await JSZip.loadAsync(await readFile(archive), {
    checkCRC32: true,
    createFolders: false,
  });
  const entries = Object.values(zip.files).sort((left, right) =>
    lexicalCompare(left.name, right.name),
  );
  if (entries.length === 0) fail("archive contains no entries.");

  const seen = new Set();
  for (const entry of entries) {
    const rawName = entry.unsafeOriginalName ?? entry.name;
    const pathname = rawName.replace(/\/$/u, "");
    if (!safeZipEntryPath(pathname)) {
      fail(`archive contains unsafe path ${JSON.stringify(rawName)}.`);
    }
    if (pathname !== expectedRoot && !pathname.startsWith(`${expectedRoot}/`)) {
      fail(`archive entry is outside ${expectedRoot}: ${rawName}.`);
    }
    if (seen.has(pathname)) fail(`archive contains duplicate path ${pathname}.`);
    seen.add(pathname);
    const mode = zipEntryMode(entry);
    if (mode != null && (mode & 0o170000) === 0o120000) {
      fail(`archive contains a symlink: ${rawName}.`);
    }
  }

  const resolvedDestination = path.resolve(destination);
  await mkdir(resolvedDestination, { recursive: true });
  for (const entry of entries) {
    if (entry.dir) continue;
    const pathname = (entry.unsafeOriginalName ?? entry.name).replace(/\/$/u, "");
    const output = path.resolve(resolvedDestination, ...pathname.split("/"));
    if (!output.startsWith(`${resolvedDestination}${path.sep}`)) {
      fail(`archive entry escaped extraction root: ${pathname}.`);
    }
    await mkdir(path.dirname(output), { recursive: true });
    await writeFile(output, await entry.async("nodebuffer"), {
      flag: "wx",
      mode: 0o644,
    });
  }
  const extractedRoot = path.join(destination, expectedRoot);
  const rootMetadata = await lstat(extractedRoot);
  if (!rootMetadata.isDirectory() || rootMetadata.isSymbolicLink()) {
    fail("archive root is not a real directory.");
  }
  return extractedRoot;
}

async function extractPinnedArchive(archive, expectedRoot, destination) {
  if (path.extname(archive).toLowerCase() === ".zip") {
    return extractPinnedZip(archive, expectedRoot, destination);
  }
  const listing = run("tar", ["-tzf", archive]).stdout;
  validateTarListing(listing, expectedRoot);
  await mkdir(destination, { recursive: true });
  run("tar", ["-xzf", archive, "-C", destination]);
  const extractedRoot = path.join(destination, expectedRoot);
  const rootMetadata = await lstat(extractedRoot);
  if (!rootMetadata.isDirectory() || rootMetadata.isSymbolicLink()) {
    fail("archive root is not a real directory.");
  }
  return extractedRoot;
}

async function assertTreeContainsNoLinks(root) {
  async function walk(directory) {
    const children = await readdir(directory, { withFileTypes: true });
    for (const child of children) {
      const absolute = path.join(directory, child.name);
      const metadata = await lstat(absolute);
      if (metadata.isSymbolicLink()) fail(`package dependency contains a symlink: ${absolute}.`);
      if (metadata.isDirectory()) await walk(absolute);
      else if (!metadata.isFile()) fail(`package dependency contains a special file: ${absolute}.`);
    }
  }
  await walk(root);
}

function findInstalledPackage(name, fromDirectory, repositoryRoot) {
  let cursor = path.resolve(fromDirectory);
  const boundary = path.resolve(repositoryRoot);
  while (cursor === boundary || cursor.startsWith(`${boundary}${path.sep}`)) {
    // Match Node's lookup order without inventing a node_modules/node_modules
    // level when the current ancestor is already the dependency directory.
    // The old path could accidentally resolve a duplicate package there and
    // copy it to an unreachable location in the standalone bundle.
    const modulesDirectory = path.basename(cursor) === "node_modules"
      ? cursor
      : path.join(cursor, "node_modules");
    const candidate = path.join(modulesDirectory, ...name.split("/"));
    if (existsSync(path.join(candidate, "package.json"))) return candidate;
    const parent = path.dirname(cursor);
    if (parent === cursor) break;
    cursor = parent;
  }
  return null;
}

async function copyPackageWithoutNestedModules(source, destination) {
  await assertTreeContainsNoLinks(source);
  await cp(source, destination, {
    recursive: true,
    force: false,
    errorOnExist: true,
    filter: (sourcePath) =>
      sourcePath === source || path.basename(sourcePath) !== "node_modules",
  });
}

async function installProductionDependencies({
  packageMetadata,
  appNodeModules,
  repositoryRoot,
}) {
  const queue = Object.keys(packageMetadata.dependencies ?? {}).map((name) => ({
    name,
    from: repositoryRoot,
    optional: false,
  }));
  const copied = new Map();

  while (queue.length > 0) {
    const request = queue.shift();
    const source = findInstalledPackage(request.name, request.from, repositoryRoot);
    if (!source) {
      if (request.optional) continue;
      fail(`production dependency ${request.name} is not installed.`);
    }
    if (copied.has(source)) continue;
    const relative = path.relative(path.join(repositoryRoot, "node_modules"), source);
    if (!safeRelativePath(relative.split(path.sep).join("/"))) {
      fail(`production dependency escaped node_modules: ${source}.`);
    }
    const metadata = JSON.parse(await readFile(path.join(source, "package.json"), "utf8"));
    if (metadata.name !== request.name) {
      fail(`resolved ${request.name} to package ${metadata.name ?? "(unnamed)"}.`);
    }
    const destination = path.join(appNodeModules, relative);
    await mkdir(path.dirname(destination), { recursive: true });
    await copyPackageWithoutNestedModules(source, destination);
    copied.set(source, {
      name: metadata.name,
      version: metadata.version,
      license: metadata.license ?? "NOASSERTION",
      source,
      relative,
    });
    for (const dependency of Object.keys(metadata.dependencies ?? {})) {
      queue.push({ name: dependency, from: source, optional: false });
    }
    for (const dependency of Object.keys(metadata.optionalDependencies ?? {})) {
      queue.push({ name: dependency, from: source, optional: true });
    }
  }

  return [...copied.values()].sort((left, right) =>
    lexicalCompare(`${left.name}@${left.version}`, `${right.name}@${right.version}`),
  );
}

async function installNativeCodecPackage({ target, appNodeModules, repositoryRoot, officeKitVersion }) {
  const name = `office-kit-codec-${target}`;
  const source = path.join(repositoryRoot, "packages", name);
  const metadata = JSON.parse(await readFile(path.join(source, "package.json"), "utf8"));
  const manifest = JSON.parse(await readFile(path.join(source, "manifest.json"), "utf8"));
  if (metadata.name !== name || metadata.version !== officeKitVersion || manifest.schemaVersion !== 2 || manifest.target !== target ||
      manifest.packageVersion !== officeKitVersion || manifest.backend !== "native-aot") {
    fail(`native codec package ${name} is missing or does not match OfficeKit ${officeKitVersion}.`);
  }
  for (const assemblyName of ["officekit-codec", "officekit-ppj-codec"]) {
    const executable = target === "win32-x64" ? `${assemblyName}.exe` : assemblyName;
    if (manifest.profiles?.[assemblyName === "officekit-codec" ? "office" : "ppj"]?.executable !== `bin/${executable}`) {
      fail(`native codec package ${name} has an invalid ${assemblyName} profile.`);
    }
    await regularFile(path.join(source, "bin", executable), `${target} ${assemblyName}`);
  }
  const destination = path.join(appNodeModules, name);
  await copyPackageWithoutNestedModules(source, destination);
  return {
    name,
    version: metadata.version,
    license: metadata.license ?? "NOASSERTION",
    source,
    relative: name,
  };
}

async function packOfficeKit({ repositoryRoot, destination }) {
  const packDirectory = path.join(destination, "npm-pack");
  await mkdir(packDirectory);
  const result = run(
    "npm",
    ["pack", "--json", "--pack-destination", packDirectory],
    { cwd: repositoryRoot },
  );
  let report;
  try {
    report = JSON.parse(result.stdout);
  } catch {
    fail("npm pack did not return JSON metadata.");
  }
  if (!Array.isArray(report) || report.length !== 1 || !report[0]?.filename) {
    fail("npm pack returned an unexpected result.");
  }
  const archive = path.join(packDirectory, report[0].filename);
  await regularFile(archive, "OfficeKit npm archive");
  const extraction = path.join(destination, "npm-extract");
  const packageRoot = await extractPinnedArchive(archive, "package", extraction);
  return { archive, packageRoot, report: report[0] };
}

function launcherScript(target) {
  if (isWindowsTarget(target)) {
    return `@echo off
setlocal EnableExtensions
set "ROOT=%~dp0.."
set "NODE=%ROOT%\\runtime\\node\\node.exe"
set "ENTRY=%ROOT%\\app\\node_modules\\office-kit\\bin\\officekit.mjs"
if not exist "%NODE%" (
  echo OfficeKit installation is incomplete: bundled Node is missing. 1>&2
  exit /b 1
)
if not exist "%ENTRY%" (
  echo OfficeKit installation is incomplete: command entrypoint is missing. 1>&2
  exit /b 1
)
"%NODE%" --max-semi-space-size=1 "%ENTRY%" %*
exit /b %ERRORLEVEL%
`;
  }
  return `#!/bin/sh
set -eu
self=$0
while [ -L "$self" ]; do
  directory=$(CDPATH= cd -- "$(dirname -- "$self")" && pwd)
  target=$(readlink "$self")
  case "$target" in
    /*) self=$target ;;
    *) self=$directory/$target ;;
  esac
done
root=$(CDPATH= cd -- "$(dirname -- "$self")/.." && pwd)
exec "$root/runtime/node/bin/node" --max-semi-space-size=1 "$root/app/node_modules/office-kit/bin/officekit.mjs" "$@"
`;
}

function licenseValue(value) {
  if (typeof value === "string" && value.trim()) return value.trim();
  if (value && typeof value === "object" && typeof value.type === "string") return value.type;
  return "NOASSERTION";
}

function cycloneDxLicense(value) {
  const normalized = licenseValue(value);
  return /[\s()]/u.test(normalized)
    ? { expression: normalized }
    : { id: normalized };
}

function componentPurl(name, version) {
  const encodedName = name.startsWith("@")
    ? `@${name.slice(1).split("/").map(encodeURIComponent).join("/")}`
    : encodeURIComponent(name);
  return `pkg:npm/${encodedName}@${encodeURIComponent(version)}`;
}

function deterministicUuid(...values) {
  const digest = sha256(Buffer.from(values.join("\0"), "utf8"));
  const variant = (Number.parseInt(digest[16], 16) & 0x3) | 0x8;
  return `urn:uuid:${digest.slice(0, 8)}-${digest.slice(8, 12)}-5${digest.slice(13, 16)}-${variant.toString(16)}${digest.slice(17, 20)}-${digest.slice(20, 32)}`;
}

function standaloneSbom({ packageMetadata, nodeVersion, target, dependencies }) {
  const components = [
    {
      type: "application",
      name: "office-kit",
      version: packageMetadata.version,
      purl: componentPurl("office-kit", packageMetadata.version),
      licenses: [{ license: cycloneDxLicense(packageMetadata.license) }],
    },
    {
      type: "framework",
      name: "Node.js",
      version: nodeVersion,
      licenses: [{ license: { name: "Node.js license" } }],
      properties: [{ name: "office-kit:target", value: target }],
    },
    ...dependencies.map((dependency) => ({
      type: "library",
      name: dependency.name,
      version: dependency.version,
      purl: componentPurl(dependency.name, dependency.version),
      licenses: [{ license: cycloneDxLicense(dependency.license) }],
    })),
  ];
  return {
    bomFormat: "CycloneDX",
    specVersion: "1.5",
    serialNumber: deterministicUuid(
      packageMetadata.version,
      nodeVersion,
      target,
      ...dependencies.map(({ name, version }) => `${name}@${version}`),
    ),
    version: 1,
    metadata: {
      component: {
        type: "application",
        name: "OfficeKit standalone",
        version: packageMetadata.version,
        properties: [{ name: "office-kit:target", value: target }],
      },
    },
    components,
  };
}

async function firstLicenseFile(packageDirectory) {
  const entries = (await readdir(packageDirectory, { withFileTypes: true }))
    .filter((entry) => entry.isFile() && /^licen[cs]e(?:[._-].*)?$/iu.test(entry.name))
    .map((entry) => entry.name)
    .sort(lexicalCompare);
  return entries[0] ? path.join(packageDirectory, entries[0]) : null;
}

async function writeLicenses({
  bundleRoot,
  repositoryRoot,
  runtimeRoot,
  dependencies,
}) {
  const licenseRoot = path.join(bundleRoot, "licenses");
  const npmLicenseRoot = path.join(licenseRoot, "npm");
  await mkdir(npmLicenseRoot, { recursive: true });
  await cp(path.join(repositoryRoot, "LICENSE"), path.join(licenseRoot, "OFFICEKIT-LICENSE.txt"));
  await cp(
    path.join(repositoryRoot, "THIRD_PARTY_NOTICES.md"),
    path.join(licenseRoot, "OFFICEKIT-THIRD-PARTY-NOTICES.md"),
  );
  await cp(path.join(runtimeRoot, "LICENSE"), path.join(licenseRoot, "NODE-LICENSE.txt"));
  for (const dependency of dependencies) {
    const safeName = `${dependency.name.replaceAll("@", "").replaceAll("/", "__")}-${
      dependency.version
    }.LICENSE.txt`;
    const source = await firstLicenseFile(dependency.source);
    if (source) {
      await cp(source, path.join(npmLicenseRoot, safeName));
    } else {
      await writeFile(
        path.join(npmLicenseRoot, safeName),
        `${dependency.name} ${dependency.version}\nDeclared license: ${licenseValue(dependency.license)}\n`,
        "utf8",
      );
    }
  }
}

async function fileInventory(root, excluded = new Set()) {
  const records = [];
  async function walk(directory, relative = "") {
    const children = await readdir(directory, { withFileTypes: true });
    children.sort((left, right) => lexicalCompare(left.name, right.name));
    for (const child of children) {
      const childRelative = relative ? `${relative}/${child.name}` : child.name;
      if (excluded.has(childRelative)) continue;
      const absolute = path.join(directory, child.name);
      const metadata = await lstat(absolute);
      if (metadata.isSymbolicLink()) fail(`bundle contains a symlink: ${childRelative}.`);
      if (metadata.isDirectory()) {
        await walk(absolute, childRelative);
      } else if (metadata.isFile()) {
        const bytes = await readFile(absolute);
        records.push({
          path: childRelative,
          bytes: bytes.length,
          sha256: sha256(bytes),
          mode: (metadata.mode & 0o111) !== 0 ? "0755" : "0644",
        });
      } else {
        fail(`bundle contains an unsupported entry: ${childRelative}.`);
      }
    }
  }
  await walk(root);
  return records;
}

function writeString(buffer, offset, length, value) {
  const bytes = Buffer.from(value, "utf8");
  if (bytes.length > length) fail(`USTAR field is too long: ${value}.`);
  bytes.copy(buffer, offset);
}

function writeOctal(buffer, offset, length, value) {
  const text = Number(value).toString(8);
  if (text.length > length - 1) fail(`USTAR numeric field is too large: ${value}.`);
  writeString(buffer, offset, length - 1, text.padStart(length - 1, "0"));
  buffer[offset + length - 1] = 0;
}

function splitUstarPath(entryPath) {
  if (Buffer.byteLength(entryPath, "utf8") <= 100) return { name: entryPath, prefix: "" };
  const segments = entryPath.split("/");
  for (let index = 1; index < segments.length; index += 1) {
    const prefix = segments.slice(0, index).join("/");
    const name = segments.slice(index).join("/");
    if (
      Buffer.byteLength(prefix, "utf8") <= 155 &&
      Buffer.byteLength(name, "utf8") <= 100
    ) {
      return { name, prefix };
    }
  }
  fail(`bundle path cannot be represented by strict USTAR: ${entryPath}.`);
}

function tarHeader({ entryPath, type, mode, size }) {
  const header = Buffer.alloc(BLOCK_SIZE);
  const { name, prefix } = splitUstarPath(entryPath);
  writeString(header, 0, 100, name);
  writeOctal(header, 100, 8, mode);
  writeOctal(header, 108, 8, 0);
  writeOctal(header, 116, 8, 0);
  writeOctal(header, 124, 12, size);
  writeOctal(header, 136, 12, 0);
  header.fill(0x20, 148, 156);
  header[156] = type === "directory" ? "5".charCodeAt(0) : "0".charCodeAt(0);
  writeString(header, 257, 6, "ustar");
  writeString(header, 263, 2, "00");
  writeString(header, 265, 32, "root");
  writeString(header, 297, 32, "root");
  writeString(header, 345, 155, prefix);
  let checksum = 0;
  for (const byte of header) checksum += byte;
  writeString(header, 148, 6, checksum.toString(8).padStart(6, "0"));
  header[154] = 0;
  header[155] = 0x20;
  return header;
}

async function archiveEntries(rootDirectory, rootName) {
  safeSegment(rootName, "archive root");
  const entries = [
    {
      path: rootName,
      type: "directory",
      mode: 0o755,
      bytes: Buffer.alloc(0),
    },
  ];
  async function walk(directory, relative = "") {
    const children = await readdir(directory, { withFileTypes: true });
    children.sort((left, right) => lexicalCompare(left.name, right.name));
    for (const child of children) {
      const relativePath = relative ? `${relative}/${child.name}` : child.name;
      if (!safeRelativePath(relativePath)) fail(`bundle contains unsafe path ${relativePath}.`);
      const absolute = path.join(directory, child.name);
      const metadata = await lstat(absolute);
      if (metadata.isSymbolicLink()) fail(`bundle contains a symlink: ${relativePath}.`);
      if (metadata.isDirectory()) {
        entries.push({
          path: `${rootName}/${relativePath}`,
          type: "directory",
          mode: 0o755,
          bytes: Buffer.alloc(0),
        });
        await walk(absolute, relativePath);
      } else if (metadata.isFile()) {
        entries.push({
          path: `${rootName}/${relativePath}`,
          type: "file",
          mode: (metadata.mode & 0o111) !== 0 ? 0o755 : 0o644,
          bytes: await readFile(absolute),
        });
      } else {
        fail(`bundle contains an unsupported entry: ${relativePath}.`);
      }
    }
  }
  await walk(rootDirectory);
  return entries;
}

export async function createDeterministicTarGz(rootDirectory, rootName) {
  const entries = await archiveEntries(rootDirectory, rootName);
  const records = [];
  for (const entry of entries) {
    records.push(
      tarHeader({
        entryPath: entry.path,
        type: entry.type,
        mode: entry.mode,
        size: entry.bytes.length,
      }),
    );
    if (entry.type === "file") {
      records.push(entry.bytes);
      const padding = (BLOCK_SIZE - (entry.bytes.length % BLOCK_SIZE)) % BLOCK_SIZE;
      if (padding) records.push(Buffer.alloc(padding));
    }
  }
  records.push(Buffer.alloc(BLOCK_SIZE * 2));
  // Node's native zlib output can change between the zlib versions bundled by
  // different Node releases. pako is pinned in package-lock.json and produces
  // the same DEFLATE stream on every supported build host.
  const archive = Buffer.from(
    pako.gzip(Buffer.concat(records), { level: 9, mtime: 0 }),
  );
  // RFC 1952 permits 255 for an unknown originating OS. Fixing this byte
  // keeps release bytes identical across the supported build hosts.
  archive[9] = 255;
  return archive;
}

export async function createDeterministicZip(rootDirectory, rootName) {
  const entries = await archiveEntries(rootDirectory, rootName);
  const zip = new JSZip();
  for (const entry of entries) {
    if (entry.type !== "file") continue;
    zip.file(entry.path, entry.bytes, {
      binary: true,
      compression: "DEFLATE",
      compressionOptions: { level: 9 },
      createFolders: false,
      date: ZIP_TIMESTAMP,
    });
  }
  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
    comment: "",
    platform: "DOS",
  });
}

async function writeBundle({
  bundleRoot,
  packageMetadata,
  nodeVersion,
  target,
  runtimeRoot,
  npmPackageRoot,
  repositoryRoot,
}) {
  const windows = isWindowsTarget(target);
  const runtimeBin = windows
    ? path.join(bundleRoot, "runtime", "node")
    : path.join(bundleRoot, "runtime", "node", "bin");
  const appNodeModules = path.join(bundleRoot, "app", "node_modules");
  const officeKitDestination = path.join(appNodeModules, "office-kit");
  await mkdir(runtimeBin, { recursive: true });
  await mkdir(appNodeModules, { recursive: true });

  const nodeFile = windows ? "node.exe" : "node";
  const nodeSource = windows
    ? path.join(runtimeRoot, nodeFile)
    : path.join(runtimeRoot, "bin", nodeFile);
  await regularFile(nodeSource, "Node executable");
  await regularFile(path.join(runtimeRoot, "LICENSE"), "Node license");
  await cp(nodeSource, path.join(runtimeBin, nodeFile));
  if (!windows) await chmod(path.join(runtimeBin, nodeFile), 0o755);
  await assertTreeContainsNoLinks(npmPackageRoot);
  await cp(npmPackageRoot, officeKitDestination, { recursive: true });
  const packedMetadata = JSON.parse(
    await readFile(path.join(officeKitDestination, "package.json"), "utf8"),
  );
  if (
    packedMetadata.name !== packageMetadata.name ||
    packedMetadata.version !== packageMetadata.version
  ) {
    fail("packed OfficeKit metadata does not match the repository package.");
  }

  const dependencies = await installProductionDependencies({
    packageMetadata,
    appNodeModules,
    repositoryRoot,
  });
  dependencies.push(await installNativeCodecPackage({
    target,
    appNodeModules,
    repositoryRoot,
    officeKitVersion: packageMetadata.version,
  }));
  dependencies.sort((left, right) => lexicalCompare(`${left.name}@${left.version}`, `${right.name}@${right.version}`));
  const binRoot = path.join(bundleRoot, "bin");
  const libRoot = path.join(bundleRoot, "lib");
  await mkdir(binRoot, { recursive: true });
  await mkdir(libRoot, { recursive: true });
  const launcher = windows ? "officekit.cmd" : "officekit";
  await writeFile(path.join(binRoot, launcher), launcherScript(target), {
    mode: windows ? 0o644 : 0o755,
  });
  if (!windows) await chmod(path.join(binRoot, launcher), 0o755);
  await cp(
    path.join(repositoryRoot, "standalone", "verify-install.mjs"),
    path.join(libRoot, "verify-install.mjs"),
  );

  await writeLicenses({
    bundleRoot,
    repositoryRoot,
    runtimeRoot,
    dependencies,
  });
  const sbom = standaloneSbom({
    packageMetadata,
    nodeVersion,
    target,
    dependencies,
  });
  await writeFile(path.join(bundleRoot, "sbom.cdx.json"), stableJson(sbom), "utf8");

  const inventory = await fileInventory(
    bundleRoot,
    new Set(["standalone-manifest.json"]),
  );
  const manifest = {
    schema: "office-kit.standalone.v1",
    officeKitVersion: packageMetadata.version,
    nodeVersion,
    target,
    entrypoint: `bin/${launcher}`,
    fileCount: inventory.length,
    unpackedBytes: inventory.reduce((total, entry) => total + entry.bytes, 0),
    files: inventory,
  };
  const manifestBytes = Buffer.from(stableJson(manifest), "utf8");
  await writeFile(
    path.join(bundleRoot, "standalone-manifest.json"),
    manifestBytes,
  );
  return { dependencies, manifest, manifestBytes: manifestBytes.length, sbom };
}

export async function buildStandalone({
  target,
  outputDirectory,
  runtimeCacheDirectory,
  runtimeArchive,
  runtimeEntry,
  nodeVersion,
  force = false,
  repositoryRoot = REPOSITORY_ROOT,
} = {}) {
  if (!SUPPORTED_TARGETS.has(target)) {
    fail(`target must be one of ${[...SUPPORTED_TARGETS].join(", ")}.`);
  }
  const packageMetadata = JSON.parse(
    await readFile(path.join(repositoryRoot, "package.json"), "utf8"),
  );
  safeSegment(packageMetadata.version, "OfficeKit version");

  const catalog = runtimeEntry
    ? null
    : await runtimeCatalog();
  const resolvedNodeVersion = nodeVersion ?? catalog?.nodeVersion;
  const resolvedEntry = runtimeEntry ?? catalog?.runtimes?.[target];
  await validateRuntimeEntry(resolvedEntry, target, resolvedNodeVersion);
  const archivePath = runtimeArchive
    ? path.resolve(runtimeArchive)
    : await downloadPinnedRuntime(
        resolvedEntry,
        runtimeCacheDirectory ??
          path.join(os.homedir(), ".cache", "office-kit", "node"),
      );
  await verifyPinnedFile(archivePath, resolvedEntry, "Node runtime archive");

  const output = path.resolve(outputDirectory);
  await mkdir(output, { recursive: true });
  const baseName = `office-kit-${packageMetadata.version}-${target}`;
  const archiveDestination = path.join(
    output,
    `${baseName}${standaloneArchiveExtension(target)}`,
  );
  const checksumDestination = `${archiveDestination}.sha256`;
  const releaseDestination = path.join(output, `${baseName}.release.json`);
  const sbomDestination = path.join(output, `${baseName}.sbom.cdx.json`);
  const noticesDestination = path.join(
    output,
    `${baseName}.THIRD_PARTY_NOTICES.md`,
  );
  for (const pathname of [
    archiveDestination,
    checksumDestination,
    releaseDestination,
    sbomDestination,
    noticesDestination,
  ]) {
    try {
      await access(pathname);
      if (!force) fail(`refusing to overwrite ${pathname}; pass --force.`);
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
    }
  }

  const temporary = await mkdtemp(path.join(output, `.standalone-${target}-`));
  try {
    const runtimeExtraction = path.join(temporary, "runtime");
    const runtimeRoot = await extractPinnedArchive(
      archivePath,
      resolvedEntry.root,
      runtimeExtraction,
    );
    const npmPack = await packOfficeKit({
      repositoryRoot,
      destination: temporary,
    });
    const bundleRoot = path.join(temporary, baseName);
    await mkdir(bundleRoot);
    const bundle = await writeBundle({
      bundleRoot,
      packageMetadata,
      nodeVersion: resolvedNodeVersion,
      target,
      runtimeRoot,
      npmPackageRoot: npmPack.packageRoot,
      repositoryRoot,
    });
    const archiveBytes = isWindowsTarget(target)
      ? await createDeterministicZip(bundleRoot, baseName)
      : await createDeterministicTarGz(bundleRoot, baseName);
    const archiveHash = sha256(archiveBytes);
    const sbomBytes = await readFile(path.join(bundleRoot, "sbom.cdx.json"));
    const noticesBytes = await readFile(
      path.join(
        bundleRoot,
        "licenses",
        "OFFICEKIT-THIRD-PARTY-NOTICES.md",
      ),
    );
    const temporaryArchive = `${archiveDestination}.tmp-${process.pid}`;
    await writeFile(temporaryArchive, archiveBytes, { flag: "wx", mode: 0o644 });
    await rename(temporaryArchive, archiveDestination);
    await writeFile(
      checksumDestination,
      `${archiveHash}  ${path.basename(archiveDestination)}\n`,
      "utf8",
    );
    await writeFile(sbomDestination, sbomBytes);
    await writeFile(noticesDestination, noticesBytes);
    const release = {
      schema: "office-kit.standalone-release.v1",
      officeKitVersion: packageMetadata.version,
      nodeVersion: resolvedNodeVersion,
      target,
      asset: path.basename(archiveDestination),
      sha256: archiveHash,
      size: archiveBytes.length,
      unpackedBytes: bundle.manifest.unpackedBytes + bundle.manifestBytes,
      fileCount: bundle.manifest.fileCount + 1,
      npmPackage: {
        filename: npmPack.report.filename,
        size: npmPack.report.size,
        unpackedSize: npmPack.report.unpackedSize,
      },
      runtime: {
        source: resolvedEntry.url,
        archive: resolvedEntry.archive,
        sha256: resolvedEntry.sha256,
        size: resolvedEntry.size,
      },
      sbom: {
        asset: path.basename(sbomDestination),
        embeddedPath: "sbom.cdx.json",
        sha256: sha256(sbomBytes),
        size: sbomBytes.length,
      },
      notices: {
        asset: path.basename(noticesDestination),
        embeddedPath: "licenses/OFFICEKIT-THIRD-PARTY-NOTICES.md",
        sha256: sha256(noticesBytes),
        size: noticesBytes.length,
      },
    };
    await writeFile(releaseDestination, stableJson(release), "utf8");
    return {
      archive: archiveDestination,
      checksum: checksumDestination,
      release: releaseDestination,
      sbom: sbomDestination,
      notices: noticesDestination,
      metadata: release,
    };
  } finally {
    await rm(temporary, { recursive: true, force: true });
  }
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const result = await buildStandalone(options);
  process.stdout.write(`${stableJson(result.metadata)}`);
}

if (
  process.argv[1] &&
  import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href
) {
  main().catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 2;
  });
}
