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
export { copyOnWriteSmartArtDefinition } from "./smartart-definition.mjs";

export async function loadPpjWorkspace(ppjPath, { cwd = process.cwd(), retainRoot = true } = {}) {
  const file = await resolveRegularFile(path.resolve(cwd, ppjPath), "PPJ input");
  if (path.extname(file).toLowerCase() !== ".ppj") throw new Error(`PPJ input must be a .ppj file: ${file}`);
  const program = await readFile(file);
  if (program.byteLength === 0 || program.byteLength > PPJ_MAX_BYTES) {
    throw new Error(`PPJ input must contain 1 through ${PPJ_MAX_BYTES} bytes: ${file}`);
  }

  const root = retainRoot ? parseProgram(program) : null;
  const resources = retainRoot ? root : parseResourceManifest(program);

  const directory = path.dirname(file);
  const assets = [];
  if (Array.isArray(resources?.assets)) {
    for (const declaration of resources.assets) {
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
  if (resources?.source && typeof resources.source === "object" && typeof resources.source.uri === "string") {
    sourcePath = await resolveWorkspaceResource(directory, resources.source.uri, "PPJ source package");
    source = await readFile(sourcePath);
  }

  return Object.freeze({
    path: file,
    directory,
    program,
    root: retainRoot ? root : null,
    source,
    sourcePath,
    assets: Object.freeze(assets),
  });
}

function parseProgram(program) {
  try {
    return JSON.parse(program.toString("utf8"));
  } catch {
    // The NativeAOT validator owns JSON diagnostics. Keeping an empty resource
    // manifest here lets it report the exact syntax path without Node inventing
    // a second language parser.
    return null;
  }
}

function parseResourceManifest(program) {
  try {
    // Preserve a BOM in the decoded text so it follows the same native
    // invalid-JSON path instead of being silently stripped by TextDecoder.
    const text = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(program);
    const parser = new ResourceManifestParser(text);
    return parser.parse();
  } catch {
    // Invalid JSON and invalid UTF-8 stay on the native validation path. Do not
    // produce competing JavaScript syntax diagnostics.
    return null;
  }
}

class ResourceManifestParser {
  #text;
  #index = 0;

  constructor(text) {
    this.#text = text;
  }

  parse() {
    const result = {};
    this.#space();
    this.#expect("{");
    this.#space();
    if (this.#take("}")) return this.#finish(result);
    while (true) {
      const keyStart = this.#index;
      this.#string();
      const key = JSON.parse(this.#text.slice(keyStart, this.#index));
      this.#space();
      this.#expect(":");
      this.#space();
      const valueStart = this.#index;
      this.#value(1);
      if (key === "assets" || key === "source") {
        result[key] = JSON.parse(this.#text.slice(valueStart, this.#index));
      }
      this.#space();
      if (this.#take("}")) return this.#finish(result);
      this.#expect(",");
      this.#space();
    }
  }

  #finish(result) {
    this.#space();
    if (this.#index !== this.#text.length) throw new SyntaxError("Trailing JSON data.");
    return result;
  }

  #value(depth) {
    if (depth > 96) throw new SyntaxError("JSON nesting exceeds the native parser budget.");
    const token = this.#text[this.#index];
    if (token === "\"") return this.#string();
    if (token === "{") return this.#object(depth + 1);
    if (token === "[") return this.#array(depth + 1);
    if (token === "t") return this.#literal("true");
    if (token === "f") return this.#literal("false");
    if (token === "n") return this.#literal("null");
    return this.#number();
  }

  #object(depth) {
    this.#expect("{");
    this.#space();
    if (this.#take("}")) return;
    while (true) {
      this.#string();
      this.#space();
      this.#expect(":");
      this.#space();
      this.#value(depth);
      this.#space();
      if (this.#take("}")) return;
      this.#expect(",");
      this.#space();
    }
  }

  #array(depth) {
    this.#expect("[");
    this.#space();
    if (this.#take("]")) return;
    while (true) {
      this.#value(depth);
      this.#space();
      if (this.#take("]")) return;
      this.#expect(",");
      this.#space();
    }
  }

  #string() {
    this.#expect("\"");
    while (this.#index < this.#text.length) {
      const code = this.#text.charCodeAt(this.#index++);
      if (code === 0x22) return;
      if (code < 0x20) throw new SyntaxError("Invalid JSON string.");
      if (code !== 0x5c) continue;
      const escape = this.#text[this.#index++];
      if ('\"\\/bfnrt'.includes(escape)) continue;
      if (escape !== "u" || !/^[0-9a-fA-F]{4}$/u.test(this.#text.slice(this.#index, this.#index + 4))) {
        throw new SyntaxError("Invalid JSON escape.");
      }
      this.#index += 4;
    }
    throw new SyntaxError("Unterminated JSON string.");
  }

  #number() {
    const start = this.#index;
    this.#take("-");
    if (this.#take("0")) {
      if (this.#digit(this.#text.charCodeAt(this.#index))) throw new SyntaxError("Invalid JSON number.");
    } else {
      if (!this.#nonzero(this.#text.charCodeAt(this.#index))) throw new SyntaxError("Invalid JSON number.");
      while (this.#digit(this.#text.charCodeAt(this.#index))) this.#index += 1;
    }
    if (this.#take(".")) {
      if (!this.#digit(this.#text.charCodeAt(this.#index))) throw new SyntaxError("Invalid JSON number.");
      while (this.#digit(this.#text.charCodeAt(this.#index))) this.#index += 1;
    }
    if (this.#text[this.#index] === "e" || this.#text[this.#index] === "E") {
      this.#index += 1;
      if (this.#text[this.#index] === "+" || this.#text[this.#index] === "-") this.#index += 1;
      if (!this.#digit(this.#text.charCodeAt(this.#index))) throw new SyntaxError("Invalid JSON number.");
      while (this.#digit(this.#text.charCodeAt(this.#index))) this.#index += 1;
    }
    if (this.#index === start) throw new SyntaxError("Expected JSON value.");
  }

  #literal(value) {
    if (!this.#text.startsWith(value, this.#index)) throw new SyntaxError("Invalid JSON literal.");
    this.#index += value.length;
  }

  #space() {
    while (true) {
      const code = this.#text.charCodeAt(this.#index);
      if (code !== 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d) return;
      this.#index += 1;
    }
  }

  #expect(value) {
    if (!this.#take(value)) throw new SyntaxError(`Expected ${value}.`);
  }

  #take(value) {
    if (this.#text[this.#index] !== value) return false;
    this.#index += 1;
    return true;
  }

  #digit(code) {
    return code >= 0x30 && code <= 0x39;
  }

  #nonzero(code) {
    return code >= 0x31 && code <= 0x39;
  }
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
