import { createHash, randomUUID } from "node:crypto";
import {
  chmod,
  link,
  lstat,
  mkdir,
  readFile,
  realpath,
  rename,
  rm,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { compilePpjToPptx, PPJ_MAX_BYTES } from "./native.mjs";

export async function loadPpjWorkspace(ppjPath, { cwd = process.cwd() } = {}) {
  const file = await resolveRegularFile(path.resolve(cwd, ppjPath), "PPJ input");
  if (path.extname(file).toLowerCase() !== ".ppj") throw new Error(`PPJ input must be a .ppj file: ${file}`);
  const program = await readFile(file);
  if (program.byteLength === 0 || program.byteLength > PPJ_MAX_BYTES) {
    throw new Error(`PPJ input must contain 1 through ${PPJ_MAX_BYTES} bytes: ${file}`);
  }

  let root;
  try {
    root = JSON.parse(program.toString("utf8"));
  } catch {
    // The NativeAOT validator owns JSON diagnostics. Keeping an empty resource
    // manifest here lets it report the exact syntax path without Node inventing
    // a second language parser.
    root = null;
  }

  const directory = path.dirname(file);
  const assets = [];
  if (Array.isArray(root?.assets)) {
    for (const declaration of root.assets) {
      if (!declaration || typeof declaration !== "object" ||
          typeof declaration.id !== "string" || typeof declaration.uri !== "string" ||
          typeof declaration.mimeType !== "string" || typeof declaration.sha256 !== "string") continue;
      const assetPath = await resolveWorkspaceResource(directory, declaration.uri, `PPJ asset ${declaration.id}`);
      assets.push(Object.freeze({
        id: declaration.id,
        fileName: path.basename(declaration.uri),
        mimeType: declaration.mimeType,
        sha256: declaration.sha256,
        data: await readFile(assetPath),
        path: assetPath,
        uri: declaration.uri,
      }));
    }
  }

  let source = new Uint8Array();
  let sourcePath = null;
  if (root?.source && typeof root.source === "object" && typeof root.source.uri === "string") {
    sourcePath = await resolveWorkspaceResource(directory, root.source.uri, "PPJ source package");
    source = await readFile(sourcePath);
  }

  return Object.freeze({
    path: file,
    directory,
    program,
    root,
    source,
    sourcePath,
    assets: Object.freeze(assets),
  });
}

export async function validatePpjWorkspace(workspace, { includeNodeMap = true } = {}) {
  return compilePpjToPptx(workspace.program, {
    source: workspace.source,
    assets: workspace.assets,
    includeNodeMap,
    validationOnly: true,
  });
}

export async function compilePpjWorkspace(workspace, { includeNodeMap = true } = {}) {
  return compilePpjToPptx(workspace.program, {
    source: workspace.source,
    assets: workspace.assets,
    includeNodeMap,
  });
}

export async function resolveRegularFile(target, label) {
  let entry;
  try {
    entry = await lstat(target);
  } catch (error) {
    if (error?.code === "ENOENT") throw new Error(`${label} does not exist: ${target}`);
    throw error;
  }
  if (entry.isSymbolicLink() || !entry.isFile()) throw new Error(`${label} must be a non-symlink regular file: ${target}`);
  return realpath(target);
}

export async function writeImmutableContent(target, data, expectedSha256) {
  const existing = await statOrNull(target);
  if (existing) {
    if (!existing.isFile() || existing.isSymbolicLink() || sha256(await readFile(target)) !== expectedSha256) {
      throw new Error(`Content-addressed PPJ asset conflicts with existing path: ${target}`);
    }
    await chmod(target, 0o444);
    return;
  }
  await mkdir(path.dirname(target), { recursive: true });
  await writeExclusiveFile(target, data, 0o444);
}

export async function writeExclusiveFile(target, data, mode = 0o644) {
  await mkdir(path.dirname(target), { recursive: true });
  const temporary = temporaryPath(target);
  try {
    await writeFile(temporary, data, { flag: "wx", mode });
    await chmod(temporary, mode);
    await link(temporary, target);
  } finally {
    await rm(temporary, { force: true });
  }
}

export async function replaceRegularFile(target, data) {
  const entry = await lstat(target);
  if (entry.isSymbolicLink() || !entry.isFile()) throw new Error(`PPJ fix target must be a non-symlink regular file: ${target}`);
  const temporary = temporaryPath(target);
  try {
    await writeFile(temporary, data, { flag: "wx", mode: entry.mode & 0o777 });
    await chmod(temporary, entry.mode & 0o777);
    await rename(temporary, target);
  } finally {
    await rm(temporary, { force: true });
  }
}

export function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

export function prettyProgram(canonical) {
  return Buffer.from(`${JSON.stringify(JSON.parse(Buffer.from(canonical).toString("utf8")), null, 2)}\n`, "utf8");
}

async function resolveWorkspaceResource(root, uri, label) {
  if (!safeRelativeUri(uri)) throw new Error(`${label} URI must stay relative to the PPJ directory: ${uri}`);
  const target = path.resolve(root, ...uri.split("/"));
  const relative = path.relative(root, target);
  if (relative === "" || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new Error(`${label} URI escapes the PPJ directory: ${uri}`);
  }
  const file = await resolveRegularFile(target, label);
  const realRelative = path.relative(await realpath(root), file);
  if (realRelative === "" || realRelative.startsWith(`..${path.sep}`) || path.isAbsolute(realRelative)) {
    throw new Error(`${label} resolves outside the PPJ directory: ${uri}`);
  }
  return file;
}

function safeRelativeUri(uri) {
  if (!uri || uri.includes("\\") || uri.includes("\0") || uri.startsWith("/")) return false;
  if (/^[A-Za-z][A-Za-z0-9+.-]*:/u.test(uri)) return false;
  return !uri.split("/").some((segment) => segment === "..");
}

async function statOrNull(target) {
  try {
    return await lstat(target);
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    throw error;
  }
}

function temporaryPath(target) {
  return path.join(path.dirname(target), `.${path.basename(target)}.${process.pid}-${randomUUID()}.tmp`);
}
