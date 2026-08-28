import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import { lstat, readFile } from "node:fs/promises";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { createRequire } from "node:module";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

import { OfficeKitCodecError } from "./office-kit-error.mjs";

export const OFFICE_KIT_NATIVE_TRANSPORT_VERSION = 2;
export const OFFICE_KIT_NATIVE_MAX_FRAME_BYTES = 128 * 1024 * 1024;

const HANDSHAKE_BYTES = 12;
const HANDSHAKE_MAGIC = Buffer.from("OKIT", "ascii");
const REQUEST_PREFIX_BYTES = 8;
const DEFAULT_IDLE_RETIRE_MS = 1_000;
const STDERR_LIMIT = 16 * 1024;
const require = createRequire(import.meta.url);
const activeClients = new Set();

const PLATFORM_PACKAGES = Object.freeze({
  "darwin-arm64": Object.freeze({
    name: "office-kit-codec-darwin-arm64",
    os: "darwin",
    cpu: "arm64",
    executable: "officekit-codec",
  }),
  "linux-x64": Object.freeze({
    name: "office-kit-codec-linux-x64",
    os: "linux",
    cpu: "x64",
    executable: "officekit-codec",
  }),
  "win32-x64": Object.freeze({
    name: "office-kit-codec-win32-x64",
    os: "win32",
    cpu: "x64",
    executable: "officekit-codec.exe",
  }),
});

let exitHookInstalled = false;

export function officeKitNativeTarget(platform = process.platform, arch = process.arch) {
  const target = `${platform}-${arch}`;
  if (!PLATFORM_PACKAGES[target]) {
    throw nativeError("runtime_unsupported_platform", `OfficeKit NativeAOT Codec does not support ${target}.`);
  }
  return target;
}

export async function loadOfficeKitNativeDescriptor({ platform = process.platform, arch = process.arch, packageJsonPath: suppliedPackageJsonPath } = {}) {
  const target = officeKitNativeTarget(platform, arch);
  const expected = PLATFORM_PACKAGES[target];
  const packageJsonPath = suppliedPackageJsonPath ?? await resolvePlatformPackageJson(expected.name);
  const packageRoot = path.dirname(packageJsonPath);
  const [packageMetadata, manifest] = await Promise.all([
    readJson(packageJsonPath, "native codec package metadata"),
    readJson(path.join(packageRoot, "manifest.json"), "native codec manifest"),
  ]);
  const rootMetadata = await readJson(new URL("../../package.json", import.meta.url), "OfficeKit package metadata");

  if (packageMetadata.name !== expected.name || packageMetadata.version !== rootMetadata.version) {
    throw nativeError("runtime_identity_mismatch", `OfficeKit requires ${expected.name}@${rootMetadata.version}.`);
  }
  if (!packageMetadata.os?.includes(expected.os) || !packageMetadata.cpu?.includes(expected.cpu)) {
    throw nativeError("runtime_identity_mismatch", `Native codec package ${expected.name} does not declare ${target}.`);
  }
  if (manifest.schemaVersion !== 1 || manifest.backend !== "native-aot" || manifest.target !== target ||
      manifest.transportVersion !== OFFICE_KIT_NATIVE_TRANSPORT_VERSION || manifest.protocolVersion !== 2 ||
      manifest.packageVersion !== rootMetadata.version) {
    throw nativeError("runtime_identity_mismatch", `Native codec manifest for ${target} is incompatible with this OfficeKit package.`);
  }

  const executablePath = path.join(packageRoot, "bin", expected.executable);
  const relativeExecutable = `bin/${expected.executable}`;
  const executableRecord = manifest.files?.find((item) => item.path === relativeExecutable);
  if (!executableRecord || !Number.isSafeInteger(executableRecord.bytes) || executableRecord.bytes <= 0 ||
      !/^[a-f0-9]{64}$/u.test(executableRecord.sha256 || "")) {
    throw nativeError("runtime_identity_mismatch", `Native codec manifest does not contain a valid ${relativeExecutable} record.`);
  }
  const descriptor = await lstat(executablePath).catch((error) => {
    if (error?.code === "ENOENT") return null;
    throw error;
  });
  if (!descriptor?.isFile() || descriptor.isSymbolicLink() || descriptor.size !== executableRecord.bytes) {
    throw nativeError("runtime_integrity_failure", `Native codec executable is missing or has the wrong size for ${target}.`);
  }
  if (platform !== "win32" && (descriptor.mode & 0o111) === 0) {
    throw nativeError("runtime_integrity_failure", `Native codec executable is not executable for ${target}.`);
  }
  const actualHash = await sha256File(executablePath);
  if (actualHash !== executableRecord.sha256) {
    throw nativeError("runtime_integrity_failure", `Native codec executable failed SHA-256 verification for ${target}.`);
  }
  return Object.freeze({
    target,
    packageName: expected.name,
    packageRoot,
    executablePath,
    assemblyName: manifest.assemblyName,
    manifest,
  });
}

export async function startOfficeKitNativeClient(options = {}) {
  const descriptor = options.descriptor ?? await loadOfficeKitNativeDescriptor(options);
  return NativeCodecClient.start(descriptor, options);
}

class NativeCodecClient {
  static async start(descriptor, { spawnProcess = spawn, idleRetireMs = DEFAULT_IDLE_RETIRE_MS } = {}) {
    const child = spawnProcess(descriptor.executablePath, ["--serve"], {
      cwd: descriptor.packageRoot,
      env: { ...process.env, DOTNET_GCConserveMemory: "9" },
      shell: false,
      stdio: ["pipe", "pipe", "pipe"],
      windowsHide: true,
    });
    const client = new NativeCodecClient(child, descriptor, idleRetireMs);
    try {
      await client.initialize();
      activeClients.add(client);
      installExitHook();
      return client;
    } catch (error) {
      client.kill();
      throw error;
    }
  }

  constructor(child, descriptor, idleRetireMs) {
    this.child = child;
    this.descriptor = descriptor;
    this.reader = new ExactReader(child.stdout);
    this.queue = Promise.resolve();
    this.pendingRequests = 0;
    this.idleRetireMs = idleRetireMs;
    this.idleTimer = undefined;
    this.retiring = false;
    this.stderr = Buffer.alloc(0);
    this.closed = false;
    this.terminated = new Promise((resolve) => {
      let settled = false;
      const finish = (value) => {
        if (settled) return;
        settled = true;
        this.closed = true;
        activeClients.delete(this);
        resolve(value);
      };
      child.once("error", (error) => finish({ error }));
      child.once("exit", (code, signal) => finish({ code, signal }));
    });
    child.stderr.on("data", (chunk) => {
      if (this.stderr.length >= STDERR_LIMIT) return;
      const bytes = Buffer.from(chunk);
      const remaining = STDERR_LIMIT - this.stderr.length;
      this.stderr = Buffer.concat([this.stderr, bytes.subarray(0, remaining)]);
    });
  }

  async initialize() {
    const handshake = await this.readOrThrow(HANDSHAKE_BYTES, "runtime_start_failed");
    if (!handshake.subarray(0, 4).equals(HANDSHAKE_MAGIC)) {
      throw nativeError("runtime_protocol_mismatch", "OfficeKit native codec returned an invalid handshake.");
    }
    const transportVersion = handshake.readUInt32BE(4);
    const protocolVersion = handshake.readUInt32BE(8);
    if (transportVersion !== OFFICE_KIT_NATIVE_TRANSPORT_VERSION || protocolVersion !== 2) {
      throw nativeError(
        "runtime_protocol_mismatch",
        `OfficeKit native codec handshake reported transport ${transportVersion} and protocol ${protocolVersion}.`,
      );
    }
    this.unref();
    this.release();
  }

  invoke(bytes, fileBytes = undefined) {
    if (!(bytes instanceof Uint8Array)) throw new TypeError("OfficeKit native codec request must be a Uint8Array.");
    if (fileBytes != null && !(fileBytes instanceof Uint8Array)) {
      throw new TypeError("OfficeKit native codec file sidecar must be a Uint8Array.");
    }
    const fileLength = fileBytes?.byteLength ?? 0;
    if (bytes.byteLength === 0 || bytes.byteLength + fileLength > OFFICE_KIT_NATIVE_MAX_FRAME_BYTES) {
      throw nativeError("request_budget_exceeded", `OfficeKit native codec request and file sidecar exceed the ${OFFICE_KIT_NATIVE_MAX_FRAME_BYTES}-byte transport budget.`);
    }
    if (!this.tryAcquire()) throw this.terminationError();
    this.pendingRequests += 1;
    if (this.pendingRequests === 1) this.ref();
    const operation = this.queue.then(
      () => this.invokeFrame(bytes, fileBytes),
      () => this.invokeFrame(bytes, fileBytes),
    );
    this.queue = operation.catch(() => {});
    return operation.finally(() => {
      this.pendingRequests -= 1;
      if (this.pendingRequests === 0) {
        this.unref();
        this.release();
      }
    });
  }

  async invokeFrame(bytes, fileBytes) {
    if (this.closed) throw this.terminationError();
    const prefix = Buffer.allocUnsafe(REQUEST_PREFIX_BYTES);
    prefix.writeUInt32BE(bytes.byteLength, 0);
    prefix.writeUInt32BE(fileBytes?.byteLength ?? 0, 4);
    await this.write(prefix);
    await this.write(bytes);
    if (fileBytes?.byteLength) await this.write(fileBytes);
    const responsePrefix = await this.readOrThrow(4, "runtime_terminated");
    const responseLength = responsePrefix.readUInt32BE(0);
    if (responseLength === 0 || responseLength > OFFICE_KIT_NATIVE_MAX_FRAME_BYTES) {
      throw nativeError("runtime_protocol_mismatch", "OfficeKit native codec returned an invalid response frame length.");
    }
    return this.readOrThrow(responseLength, "runtime_terminated");
  }

  async write(bytes) {
    if (this.closed || this.child.stdin.destroyed) throw this.terminationError();
    if (!this.child.stdin.write(bytes)) {
      await Promise.race([
        once(this.child.stdin, "drain"),
        this.terminated.then(() => { throw this.terminationError(); }),
      ]);
    }
  }

  async readOrThrow(length, code) {
    try {
      return await Promise.race([
        this.reader.readExactly(length),
        this.terminated.then(() => { throw this.terminationError(code); }),
      ]);
    } catch (error) {
      if (error instanceof OfficeKitCodecError) throw error;
      throw nativeError(code, "OfficeKit native codec closed its response stream unexpectedly.", error);
    }
  }

  terminationError(code = "runtime_terminated") {
    const detail = this.stderr.toString("utf8").trim();
    return nativeError(code, `OfficeKit native codec terminated unexpectedly${detail ? `: ${detail}` : "."}`);
  }

  ref() {
    safeLifecycleCall(this.child, "ref");
    safeLifecycleCall(this.child.stdin, "ref");
    safeLifecycleCall(this.child.stdout, "ref");
    safeLifecycleCall(this.child.stderr, "ref");
  }

  unref() {
    safeLifecycleCall(this.child, "unref");
    safeLifecycleCall(this.child.stdin, "unref");
    safeLifecycleCall(this.child.stdout, "unref");
    safeLifecycleCall(this.child.stderr, "unref");
  }

  kill() {
    this.retiring = true;
    this.clearIdleRetirement();
    if (!this.child.killed) this.child.kill();
  }

  get idle() {
    return this.pendingRequests === 0;
  }

  tryAcquire() {
    if (this.closed || this.retiring) return false;
    this.clearIdleRetirement();
    return true;
  }

  release() {
    if (!this.idle || this.closed || this.retiring || !(this.idleRetireMs >= 0)) return;
    this.clearIdleRetirement();
    this.idleTimer = setTimeout(() => {
      this.idleTimer = undefined;
      void this.retire().catch(() => {});
    }, this.idleRetireMs);
    this.idleTimer.unref?.();
  }

  clearIdleRetirement() {
    if (this.idleTimer !== undefined) clearTimeout(this.idleTimer);
    this.idleTimer = undefined;
  }

  async retire() {
    if (this.closed) return;
    this.ref();
    this.kill();
    await this.terminated;
  }
}

class ExactReader {
  constructor(stream) {
    this.iterator = stream[Symbol.asyncIterator]();
    this.pending = null;
    this.offset = 0;
  }

  async readExactly(length) {
    const output = Buffer.allocUnsafe(length);
    let written = 0;
    while (written < length) {
      if (!this.pending || this.offset >= this.pending.length) {
        const next = await this.iterator.next();
        if (next.done) throw new Error("EOF");
        this.pending = Buffer.from(next.value.buffer, next.value.byteOffset, next.value.byteLength);
        this.offset = 0;
      }
      const available = this.pending.length - this.offset;
      const copied = Math.min(available, length - written);
      this.pending.copy(output, written, this.offset, this.offset + copied);
      written += copied;
      this.offset += copied;
    }
    return output;
  }
}

async function resolvePlatformPackageJson(packageName) {
  try {
    return require.resolve(`${packageName}/package.json`);
  } catch (error) {
    if (error?.code !== "MODULE_NOT_FOUND") throw error;
  }
  const developmentPath = path.resolve(fileURLToPath(new URL(`../../packages/${packageName}/package.json`, import.meta.url)));
  const descriptor = await lstat(developmentPath).catch((error) => error?.code === "ENOENT" ? null : Promise.reject(error));
  if (descriptor?.isFile()) return developmentPath;
  throw nativeError("runtime_unavailable", `Required native codec package ${packageName} is not installed.`);
}

async function readJson(location, label) {
  try {
    return JSON.parse(await readFile(location, "utf8"));
  } catch (error) {
    throw nativeError("runtime_identity_mismatch", `OfficeKit could not read ${label}.`, error);
  }
}

async function sha256File(file) {
  const hash = createHash("sha256");
  for await (const chunk of createReadStream(file)) hash.update(chunk);
  return hash.digest("hex");
}

function installExitHook() {
  if (exitHookInstalled) return;
  exitHookInstalled = true;
  process.once("exit", () => {
    for (const client of activeClients) client.kill();
  });
}

function safeLifecycleCall(target, method) {
  try {
    target?.[method]?.();
  } catch {
    // A child may close between the frame terminal state and the idle unref.
  }
}

function nativeError(code, message, cause) {
  return new OfficeKitCodecError(message, [], { code, cause });
}
