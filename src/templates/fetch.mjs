import { createHash, randomUUID } from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { queryTemplates } from "./search.mjs";

const HASH = /^[a-f0-9]{64}$/u;
const TEMPLATE_ID = /^artifact-template-[a-z0-9]+(?:-[a-z0-9]+)*$/u;
const MAX_PPJ_BYTES = 16 * 1024 * 1024;
const MAX_PPTX_BYTES = 256 * 1024 * 1024;
const MAX_ASSET_BYTES = 64 * 1024 * 1024;
const MAX_TOTAL_BYTES = 256 * 1024 * 1024;
const REMOTE_HOST = "raw.githubusercontent.com";
const FETCH_TIMEOUT_MS = 120_000;

export const TEMPLATE_FETCH_USAGE = [
  "Usage: officekit template fetch <artifact-template-id>",
  "  [--root <absolute-template-root>]... [--cache-root <absolute-dir>] [--json]",
].join("\n");

export async function fetchTemplateReferences({
  id,
  roots = null,
  cacheRoot = null,
  projectPath = process.cwd(),
} = {}) {
  if (!TEMPLATE_ID.test(id ?? "")) throw new Error("template fetch requires an artifact-template-* id");
  const result = await queryTemplates({
    kind: "presentation",
    id,
    roots,
    maxCandidates: 1,
    projectPath,
  });
  const candidate = result.candidates[0];
  if (candidate == null) throw new Error(`Presentation template ${id} was not found.`);
  const declarations = [candidate.referenceProgram, candidate.referencePptx].filter(Boolean);
  if (declarations.length === 0) {
    return Object.freeze({
      ok: true,
      schema: "office-kit/template-fetch/v1",
      id,
      status: "no-references",
      references: [],
    });
  }

  const officeKitHome = process.env.OFFICE_KIT_HOME == null
    ? path.join(os.homedir(), ".office-kit")
    : path.resolve(process.env.OFFICE_KIT_HOME);
  const base = path.resolve(cacheRoot ?? path.join(officeKitHome, "cache", "templates"));
  const key = candidate.referenceProgram?.sha256 ?? candidate.referencePptx?.sha256;
  const destinationRoot = path.join(base, id, key);
  await mkdirReal(destinationRoot);
  const referencesRoot = path.join(destinationRoot, "assets", "references");
  await mkdirReal(referencesRoot);
  let totalBytes = 0;
  const references = [];
  for (const declaration of declarations) {
    const target = path.join(destinationRoot, declaration.path);
    const fetched = await materializeReference(declaration, target, {
      maxBytes: path.extname(declaration.path).toLowerCase() === ".ppj" ? MAX_PPJ_BYTES : MAX_PPTX_BYTES,
      label: declaration.path,
      total: () => totalBytes,
      add: (bytes) => { totalBytes += bytes; },
    });
    references.push(fetched);
  }

  const manifest = {
    schema: "office-kit/template-fetch/v1",
    templateId: id,
    templateSchemaVersion: candidate.templateSchemaVersion,
    fetchedAt: new Date().toISOString(),
    references,
  };
  const manifestPath = path.join(destinationRoot, "template-fetch.json");
  await writeImmutable(manifestPath, Buffer.from(`${JSON.stringify(manifest, null, 2)}\n`));
  return Object.freeze({
    ok: true,
    schema: manifest.schema,
    id,
    status: "fetched",
    cacheRoot: destinationRoot,
    manifestPath,
    references,
  });
}

async function materializeReference(declaration, target, { maxBytes, label, total, add }) {
  const existing = await existingFile(target);
  if (existing != null) {
    const actual = sha256(existing.bytes);
    if (actual !== declaration.sha256) throw new Error(`${label} cache path has a SHA-256 conflict`);
    return Object.freeze({
      path: existing.path,
      sha256: declaration.sha256,
      bytes: existing.bytes.byteLength,
      reused: true,
      dependencies: await readAndFetchPpjDependencies(declaration, existing.path, null, { total, add }),
    });
  }
  if (declaration.download == null) {
    throw new Error(`${label} is not cached and has no GitHub download descriptor`);
  }
  const bytes = await fetchBytes(declaration.download.url, maxBytes, label);
  if (sha256(bytes) !== declaration.sha256) throw new Error(`${label} download SHA-256 mismatch`);
  if (total() + bytes.byteLength > MAX_TOTAL_BYTES) throw new Error("template reference downloads exceed the total byte budget");
  await writeImmutable(target, bytes);
  add(bytes.byteLength);
  const dependencies = await readAndFetchPpjDependencies(declaration, target, bytes, { total, add });
  return Object.freeze({
    path: target,
    sha256: declaration.sha256,
    bytes: bytes.byteLength,
    reused: false,
    dependencies,
  });
}

async function readAndFetchPpjDependencies(declaration, ppjPath, ppjBytes, { total, add }) {
  if (path.extname(declaration.path).toLowerCase() !== ".ppj") return [];
  const bytes = ppjBytes ?? await fs.readFile(ppjPath);
  let program;
  try { program = JSON.parse(bytes.toString("utf8")); } catch (error) {
    throw new Error(`${declaration.path} is not strict JSON: ${error.message}`);
  }
  if (program?.schema !== "office-kit/ppj/v1") throw new Error(`${declaration.path} has an unsupported PPJ schema`);
  const dependencies = [
    ...(program.source == null ? [] : [{ ...program.source, kind: "source" }]),
    ...(program.assets ?? []).map((asset) => ({ ...asset, kind: "asset" })),
  ];
  const seen = new Set();
  const output = [];
  for (const dependency of dependencies) {
    const uri = dependency.uri;
    if (!safeRelativeUri(uri) || seen.has(uri)) throw new Error(`${declaration.path} contains an unsafe or duplicate dependency URI: ${uri}`);
    seen.add(uri);
    if (!HASH.test(dependency.sha256 ?? "")) throw new Error(`${declaration.path} dependency ${uri} has no valid SHA-256`);
    const target = path.resolve(path.dirname(ppjPath), ...uri.split("/"));
    const maxBytes = dependency.kind === "source" ? MAX_PPTX_BYTES : MAX_ASSET_BYTES;
    const local = await existingFile(target);
    if (local != null) {
      if (sha256(local.bytes) !== dependency.sha256) throw new Error(`${declaration.path} dependency ${uri} cache path has a SHA-256 conflict`);
      output.push({ uri, path: local.path, sha256: dependency.sha256, bytes: local.bytes.byteLength, reused: true });
      continue;
    }
    const url = new URL(uri, declaration.download.url).toString();
    const dependencyBytes = await fetchBytes(url, maxBytes, `${declaration.path} dependency ${uri}`);
    if (sha256(dependencyBytes) !== dependency.sha256) throw new Error(`${declaration.path} dependency ${uri} download SHA-256 mismatch`);
    if (total() + dependencyBytes.byteLength > MAX_TOTAL_BYTES) throw new Error("template reference downloads exceed the total byte budget");
    await writeImmutable(target, dependencyBytes);
    add(dependencyBytes.byteLength);
    output.push({ uri, path: target, sha256: dependency.sha256, bytes: dependencyBytes.byteLength, reused: false });
  }
  return output;
}

async function fetchBytes(rawUrl, maxBytes, label) {
  const url = validateDownloadUrl(rawUrl, label);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
  try {
    const response = await fetch(url, { redirect: "error", signal: controller.signal });
    if (!response.ok) throw new Error(`${label} download failed with HTTP ${response.status}`);
    const declaredLength = response.headers.get("content-length");
    if (declaredLength != null && Number(declaredLength) > maxBytes) throw new Error(`${label} exceeds the ${maxBytes}-byte download limit`);
    if (response.body == null) throw new Error(`${label} download returned no body`);
    const reader = response.body.getReader();
    const chunks = [];
    let size = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > maxBytes) {
        await reader.cancel();
        throw new Error(`${label} exceeds the ${maxBytes}-byte download limit`);
      }
      chunks.push(Buffer.from(value));
    }
    return Buffer.concat(chunks, size);
  } catch (error) {
    if (error?.name === "AbortError") throw new Error(`${label} download timed out`);
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

function validateDownloadUrl(rawUrl, label) {
  let url;
  try { url = new URL(rawUrl); } catch { throw new Error(`${label} download URL is invalid`); }
  if (url.protocol !== "https:" || url.hostname.toLowerCase() !== REMOTE_HOST ||
      url.username || url.password || url.search || url.hash || url.pathname.includes("..")) {
    throw new Error(`${label} download URL must be an HTTPS raw.githubusercontent.com URL without credentials, query, or traversal`);
  }
  return url.toString();
}

function safeRelativeUri(uri) {
  return typeof uri === "string" && uri.length > 0 && !uri.includes("\\") && !uri.includes("\0") &&
    !uri.startsWith("/") && !/^[A-Za-z][A-Za-z0-9+.-]*:/u.test(uri) &&
    !uri.split("/").some((segment) => segment === "" || segment === "." || segment === "..");
}

async function mkdirReal(directory) {
  await fs.mkdir(directory, { recursive: true, mode: 0o755 });
  const stat = await fs.lstat(directory);
  if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error(`template cache path must be a real directory: ${directory}`);
}

async function existingFile(filePath) {
  try {
    const stat = await fs.lstat(filePath);
    if (stat.isSymbolicLink() || !stat.isFile()) throw new Error(`template cache path must be a regular file: ${filePath}`);
    return { path: await fs.realpath(filePath), bytes: await fs.readFile(filePath) };
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    throw error;
  }
}

async function writeImmutable(filePath, bytes) {
  await mkdirReal(path.dirname(filePath));
  const existing = await existingFile(filePath);
  if (existing != null) {
    if (sha256(existing.bytes) !== sha256(bytes)) throw new Error(`template cache path has a SHA-256 conflict: ${filePath}`);
    await fs.chmod(filePath, 0o444);
    return;
  }
  const temporary = `${filePath}.${process.pid}-${randomUUID()}.tmp`;
  try {
    await fs.writeFile(temporary, bytes, { flag: "wx", mode: 0o444 });
    await fs.chmod(temporary, 0o444);
    await fs.link(temporary, filePath);
  } finally {
    await fs.rm(temporary, { force: true });
  }
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

export function parseTemplateFetchArguments(args) {
  const positional = [];
  const roots = [];
  let cacheRoot = null;
  let json = false;
  let help = false;
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--json") json = true;
    else if (argument === "--help" || argument === "-h") help = true;
    else if (argument === "--root") roots.push(requiredValue(args, ++index, argument));
    else if (argument.startsWith("--root=")) roots.push(argument.slice("--root=".length));
    else if (argument === "--cache-root") cacheRoot = requiredValue(args, ++index, argument);
    else if (argument.startsWith("--cache-root=")) cacheRoot = argument.slice("--cache-root=".length);
    else if (argument.startsWith("-")) throw new Error(TEMPLATE_FETCH_USAGE);
    else positional.push(argument);
  }
  if (help) return { help, json };
  if (positional.length !== 1) throw new Error(TEMPLATE_FETCH_USAGE);
  return {
    help,
    json,
    id: positional[0],
    roots: roots.length === 0 ? null : roots.map((entry) => path.resolve(entry)),
    cacheRoot: cacheRoot == null ? null : path.resolve(cacheRoot),
  };
}

function requiredValue(args, index, option) {
  const value = args[index];
  if (value == null || value.startsWith("-")) throw new Error(`${option} requires a value`);
  return value;
}

export function formatTemplateFetchResult(result) {
  if (result.status === "no-references") return `${result.id}: no remote presentation references declared`;
  return [
    `Fetched ${result.id}`,
    `Cache     ${result.cacheRoot}`,
    ...result.references.map((reference) => `${reference.path} (${reference.bytes} bytes${reference.reused ? ", reused" : ""})`),
    `Manifest  ${result.manifestPath}`,
  ].join("\n");
}
