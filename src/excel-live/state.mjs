import { randomBytes } from "node:crypto";
import { chmod, lstat, mkdir, readFile, rename, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { excelLiveError } from "./errors.mjs";

export const EXCEL_STATE_SCHEMA = 1;
export const EXCEL_BRIDGE_PORT = 47213;
export const EXCEL_ADDIN_ID = "d209533c-4ca9-4aa1-b64b-467bbdd23fc0";

export function resolveExcelStatePaths({ env = process.env, home = os.homedir() } = {}) {
  const configuredHome = env.OFFICEKIT_EXCEL_HOME;
  const root = path.resolve(
    configuredHome && configuredHome.length > 0
      ? configuredHome
      : path.join(env.OFFICE_KIT_HOME || path.join(home, ".office-kit"), "excel"),
  );
  return Object.freeze({
    root,
    config: path.join(root, "config.json"),
    secret: path.join(root, "cli-secret"),
    rootCertificate: path.join(root, "officekit-local-root.pem"),
    rootKey: path.join(root, "officekit-local-root-key.pem"),
    leafCertificate: path.join(root, "localhost.pem"),
    leafKey: path.join(root, "localhost-key.pem"),
    manifest: path.join(root, "officekit-excel-manifest.xml"),
    audit: path.join(root, "audit.ndjson"),
    pid: path.join(root, "bridge.pid"),
  });
}

export async function ensureExcelStateDirectory(paths) {
  await mkdir(paths.root, { recursive: true, mode: 0o700 });
  await assertRegularDirectory(paths.root, "Excel state directory");
  if (process.platform !== "win32") await chmod(paths.root, 0o700);
}

export async function readExcelConfiguration(paths) {
  const config = await readJsonFile(paths.config, "Excel configuration");
  validateConfiguration(config);
  const secret = await readPrivateText(paths.secret, "Excel CLI secret");
  if (!/^[a-f0-9]{64}$/u.test(secret)) {
    throw excelLiveError("invalid-state", "Excel CLI secret has an invalid format.");
  }
  return { config, secret };
}

export async function initializeExcelConfiguration(paths, { port = EXCEL_BRIDGE_PORT } = {}) {
  await ensureExcelStateDirectory(paths);
  const existing = await lstatIfExists(paths.config);
  if (existing != null) return readExcelConfiguration(paths);
  if (!Number.isSafeInteger(port) || port < 1024 || port > 65535) {
    throw excelLiveError("invalid-state", "Excel bridge port must be a user port.");
  }
  const config = {
    schemaVersion: EXCEL_STATE_SCHEMA,
    addinId: EXCEL_ADDIN_ID,
    port,
    createdAt: new Date().toISOString(),
    certificate: null,
    trusted: false,
  };
  const secret = randomBytes(32).toString("hex");
  await writePrivateText(paths.secret, secret);
  await writeJsonFile(paths.config, config);
  return { config, secret };
}

export async function updateExcelConfiguration(paths, update) {
  const { config, secret } = await readExcelConfiguration(paths);
  const next = typeof update === "function" ? update(structuredClone(config)) : update;
  validateConfiguration(next);
  await writeJsonFile(paths.config, next);
  return { config: next, secret };
}

export async function writeJsonFile(target, value) {
  await writePrivateText(target, `${JSON.stringify(value, null, 2)}\n`);
}

export async function readJsonFile(target, label) {
  const content = await readPrivateText(target, label);
  try {
    return JSON.parse(content);
  } catch (error) {
    throw excelLiveError("invalid-state", `${label} is not valid JSON: ${error.message}`);
  }
}

export async function writePrivateText(target, content) {
  const parent = path.dirname(target);
  await mkdir(parent, { recursive: true, mode: 0o700 });
  await assertRegularDirectory(parent, "Excel state parent directory");
  const existing = await lstatIfExists(target);
  if (existing?.isSymbolicLink()) {
    throw excelLiveError("unsafe-state-path", `Refusing to write symbolic link: ${target}`);
  }
  const temporary = `${target}.${process.pid}-${randomBytes(8).toString("hex")}.tmp`;
  await writeFile(temporary, content, { encoding: "utf8", flag: "wx", mode: 0o600 });
  if (process.platform !== "win32") await chmod(temporary, 0o600);
  try {
    await rename(temporary, target);
  } finally {
    await rm(temporary, { force: true });
  }
}

export async function readPrivateText(target, label) {
  const stat = await lstatIfExists(target);
  if (stat == null) throw excelLiveError("not-installed", `${label} is missing. Run officekit excel install.`);
  if (stat.isSymbolicLink() || !stat.isFile()) {
    throw excelLiveError("unsafe-state-path", `${label} must be a regular file.`);
  }
  if (stat.size > 2_000_000) throw excelLiveError("invalid-state", `${label} is too large.`);
  return (await readFile(target, "utf8")).trim();
}

export async function removeExcelState(paths) {
  const stat = await lstatIfExists(paths.root);
  if (stat == null) return false;
  if (stat.isSymbolicLink() || !stat.isDirectory()) {
    throw excelLiveError("unsafe-state-path", "Excel state directory must be a regular directory.");
  }
  await rm(paths.root, { recursive: true, force: false, maxRetries: 2 });
  return true;
}

export async function appendAuditRecord(paths, record) {
  await ensureExcelStateDirectory(paths);
  const existing = await lstatIfExists(paths.audit);
  if (existing?.isSymbolicLink() || (existing != null && !existing.isFile())) {
    throw excelLiveError("unsafe-state-path", "Excel audit file must be a regular file.");
  }
  const line = `${JSON.stringify(record)}\n`;
  await writeFile(paths.audit, line, {
    encoding: "utf8",
    flag: "a",
    mode: 0o600,
  });
  if (process.platform !== "win32") await chmod(paths.audit, 0o600);
}

export async function lstatIfExists(target) {
  try {
    return await lstat(target);
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    throw error;
  }
}

async function assertRegularDirectory(target, label) {
  const stat = await lstat(target);
  if (stat.isSymbolicLink() || !stat.isDirectory()) {
    throw excelLiveError("unsafe-state-path", `${label} must be a regular directory.`);
  }
}

function validateConfiguration(config) {
  if (config == null || typeof config !== "object" || Array.isArray(config)) {
    throw excelLiveError("invalid-state", "Excel configuration must be an object.");
  }
  if (config.schemaVersion !== EXCEL_STATE_SCHEMA || config.addinId !== EXCEL_ADDIN_ID) {
    throw excelLiveError("invalid-state", "Excel configuration schema is unsupported.");
  }
  if (!Number.isSafeInteger(config.port) || config.port < 1024 || config.port > 65535) {
    throw excelLiveError("invalid-state", "Excel configuration has an invalid port.");
  }
  if (typeof config.createdAt !== "string") {
    throw excelLiveError("invalid-state", "Excel configuration has no creation timestamp.");
  }
  if (typeof config.trusted !== "boolean") {
    throw excelLiveError("invalid-state", "Excel configuration has invalid trust state.");
  }
  if (config.certificate !== null) {
    const certificate = config.certificate;
    if (
      certificate == null || typeof certificate !== "object" ||
      !/^[A-F0-9:]{59}$/u.test(certificate.rootFingerprintSha1 ?? "") ||
      !/^[A-F0-9:]{95}$/u.test(certificate.rootFingerprintSha256 ?? "") ||
      !/^[A-F0-9:]{95}$/u.test(certificate.leafFingerprintSha256 ?? "")
    ) {
      throw excelLiveError("invalid-state", "Excel configuration has invalid certificate metadata.");
    }
  }
}
