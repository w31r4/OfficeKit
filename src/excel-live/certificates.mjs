import { X509Certificate } from "node:crypto";
import { spawn } from "node:child_process";
import { rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import selfsigned from "selfsigned";

import { excelLiveError } from "./errors.mjs";
import { lstatIfExists, readPrivateText, updateExcelConfiguration, writePrivateText } from "./state.mjs";

const ROOT_COMMON_NAME = "OfficeKit Local Root";
const LEAF_COMMON_NAME = "localhost";

export async function ensureExcelCertificates(paths, config) {
  if (config.certificate != null && await certificateFilesExist(paths)) {
    // Early development builds wrote this key even though the running bridge
    // only needs the localhost leaf. Remove it on the next normal command.
    await rm(paths.rootKey, { force: true });
    return readCertificateBundle(paths, config.certificate);
  }
  const generated = await generateCertificateBundle();
  await writePrivateText(paths.rootCertificate, generated.root.cert);
  await writePrivateText(paths.leafCertificate, generated.leaf.cert);
  await writePrivateText(paths.leafKey, generated.leaf.private);
  // The root key is required only while the leaf is minted. Keeping it on
  // disk after trusting the root would unnecessarily widen the impact of a
  // compromise of this local state directory. Remove a key left by an older
  // pre-release installation too; unlinking a symlink removes the link, not
  // its target.
  await rm(paths.rootKey, { force: true });
  const root = new X509Certificate(generated.root.cert);
  const leaf = new X509Certificate(generated.leaf.cert);
  assertCertificateBundle(root, leaf);
  const certificate = {
    rootFingerprintSha1: root.fingerprint,
    rootFingerprintSha256: root.fingerprint256,
    leafFingerprintSha256: leaf.fingerprint256,
    createdAt: new Date().toISOString(),
  };
  return { ...generated, root, leaf, chain: `${generated.leaf.cert}\n${generated.root.cert}`, certificate };
}

export async function persistCertificateMetadata(paths, certificate) {
  const result = await updateExcelConfiguration(paths, (config) => ({
    ...config,
    certificate,
    trusted: false,
  }));
  return result.config;
}

export async function readCertificateBundle(paths, metadata) {
  const [rootPem, leafPem, leafKey] = await Promise.all([
    readPrivateText(paths.rootCertificate, "Excel root certificate"),
    readPrivateText(paths.leafCertificate, "Excel leaf certificate"),
    readPrivateText(paths.leafKey, "Excel leaf private key"),
  ]);
  const root = new X509Certificate(rootPem);
  const leaf = new X509Certificate(leafPem);
  assertCertificateBundle(root, leaf);
  if (
    root.fingerprint256 !== metadata.rootFingerprintSha256 ||
    leaf.fingerprint256 !== metadata.leafFingerprintSha256
  ) {
    throw excelLiveError("invalid-state", "Excel certificate files do not match saved metadata.");
  }
  return {
    root: { cert: rootPem, private: null },
    leaf: { cert: leafPem, private: leafKey },
    chain: `${leafPem}\n${rootPem}`,
    rootCertificate: root,
    leafCertificate: leaf,
    certificate: metadata,
  };
}

export async function trustExcelRootCertificate(paths, config, { run = runCommand } = {}) {
  if (config.certificate == null) {
    throw excelLiveError("not-installed", "Excel certificates are missing. Run officekit excel install.");
  }
  const command = trustCommand(paths.rootCertificate);
  await run(command.file, command.args);
  const { config: next } = await updateExcelConfiguration(paths, (current) => ({
    ...current,
    trusted: true,
  }));
  return next;
}

export async function untrustExcelRootCertificate(paths, config, { run = runCommand } = {}) {
  if (config.certificate == null || !config.trusted) return config;
  const command = untrustCommand(config.certificate.rootFingerprintSha1);
  await run(command.file, command.args);
  const { config: next } = await updateExcelConfiguration(paths, (current) => ({
    ...current,
    trusted: false,
  }));
  return next;
}

export function trustInstructions(paths) {
  if (process.platform === "darwin") {
    return `Trust ${paths.rootCertificate} in your login keychain when macOS asks.`;
  }
  if (process.platform === "win32") {
    return `Trust ${paths.rootCertificate} in the Current User Root certificate store when Windows asks.`;
  }
  return "Excel Live Control currently supports certificate trust on Windows and macOS desktop only.";
}

export async function probeExcelRootTrust(paths, config, { run = runCommand } = {}) {
  if (config.certificate == null) return { trusted: false, reason: "certificate-missing" };
  try {
    const command = probeTrustCommand(config.certificate.rootFingerprintSha1);
    const result = await run(command.file, command.args, { allowFailure: true });
    return {
      trusted: result.status === 0 && fingerprintMatches(result.stdout, config.certificate.rootFingerprintSha1),
      reason: result.status === 0 ? undefined : "root-not-in-user-store",
    };
  } catch (error) {
    return { trusted: false, reason: error.code ?? "trust-probe-failed" };
  }
}

export async function generateCertificateBundle(now = new Date()) {
  const notBeforeDate = new Date(now.getTime() - 5 * 60_000);
  const rootNotAfter = new Date(now);
  rootNotAfter.setFullYear(rootNotAfter.getFullYear() + 10);
  const leafNotAfter = new Date(now);
  leafNotAfter.setFullYear(leafNotAfter.getFullYear() + 2);
  const root = await selfsigned.generate(
    [{ name: "commonName", value: ROOT_COMMON_NAME }],
    {
      keyType: "ec",
      curve: "P-256",
      algorithm: "sha256",
      notBeforeDate,
      notAfterDate: rootNotAfter,
      extensions: [
        { name: "basicConstraints", cA: true, pathLenConstraint: 0, critical: true },
        {
          name: "keyUsage",
          digitalSignature: true,
          keyCertSign: true,
          cRLSign: true,
          critical: true,
        },
      ],
    },
  );
  const leaf = await selfsigned.generate(
    [{ name: "commonName", value: LEAF_COMMON_NAME }],
    {
      keyType: "ec",
      curve: "P-256",
      algorithm: "sha256",
      notBeforeDate,
      notAfterDate: leafNotAfter,
      ca: { key: root.private, cert: root.cert },
      extensions: [
        { name: "basicConstraints", cA: false, critical: true },
        { name: "keyUsage", digitalSignature: true, keyEncipherment: true, critical: true },
        { name: "extKeyUsage", serverAuth: true },
        {
          name: "subjectAltName",
          altNames: [
            { type: 2, value: "localhost" },
            { type: 7, ip: "127.0.0.1" },
            { type: 7, ip: "::1" },
          ],
        },
      ],
    },
  );
  return { root, leaf };
}

async function certificateFilesExist(paths) {
  for (const target of [paths.rootCertificate, paths.leafCertificate, paths.leafKey]) {
    const stat = await lstatIfExists(target);
    if (stat == null || stat.isSymbolicLink() || !stat.isFile()) return false;
  }
  return true;
}

function assertCertificateBundle(root, leaf) {
  if (root.subject !== "CN=OfficeKit Local Root") {
    throw excelLiveError("invalid-state", "Excel root certificate has an unexpected subject.");
  }
  if (leaf.subject !== "CN=localhost" || leaf.checkHost("localhost") == null || leaf.checkIP("127.0.0.1") == null) {
    throw excelLiveError("invalid-state", "Excel leaf certificate does not cover localhost.");
  }
  if (!leaf.verify(root.publicKey)) {
    throw excelLiveError("invalid-state", "Excel leaf certificate is not signed by the local root.");
  }
}

function trustCommand(rootCertificate) {
  if (process.platform === "darwin") {
    return {
      file: "security",
      args: [
        "add-trusted-cert",
        "-r",
        "trustRoot",
        "-k",
        path.join(os.homedir(), "Library", "Keychains", "login.keychain-db"),
        rootCertificate,
      ],
    };
  }
  if (process.platform === "win32") {
    return { file: "certutil.exe", args: ["-user", "-addstore", "Root", rootCertificate] };
  }
  throw excelLiveError("unsupported-platform", "Excel Live Control supports certificate trust only on Windows and macOS desktop.");
}

function untrustCommand(fingerprint) {
  const normalized = fingerprint.replaceAll(":", "");
  if (process.platform === "darwin") {
    return {
      file: "security",
      args: [
        "delete-certificate",
        "-Z",
        normalized,
        path.join(os.homedir(), "Library", "Keychains", "login.keychain-db"),
      ],
    };
  }
  if (process.platform === "win32") {
    return { file: "certutil.exe", args: ["-user", "-delstore", "Root", normalized] };
  }
  throw excelLiveError("unsupported-platform", "Excel Live Control supports certificate trust only on Windows and macOS desktop.");
}

function probeTrustCommand(fingerprint) {
  const normalized = fingerprint.replaceAll(":", "");
  if (process.platform === "darwin") {
    return {
      file: "security",
      args: [
        "find-certificate",
        "-a",
        "-Z",
        "-c",
        ROOT_COMMON_NAME,
        path.join(os.homedir(), "Library", "Keychains", "login.keychain-db"),
      ],
    };
  }
  if (process.platform === "win32") {
    return { file: "certutil.exe", args: ["-user", "-store", "Root", normalized] };
  }
  throw excelLiveError("unsupported-platform", "Excel Live Control supports certificate trust only on Windows and macOS desktop.");
}

function fingerprintMatches(output, expected) {
  const normalized = output.replace(/[^A-Fa-f0-9]/gu, "").toUpperCase();
  return normalized.includes(expected.replaceAll(":", "").toUpperCase());
}

function runCommand(file, args, { allowFailure = false } = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(file, args, { stdio: ["ignore", "pipe", "pipe"], shell: false, windowsHide: true });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => { stdout = `${stdout}${chunk}`.slice(-16_384); });
    child.stderr.on("data", (chunk) => { stderr = `${stderr}${chunk}`.slice(-16_384); });
    child.once("error", (error) => {
      reject(excelLiveError("certificate-command-failed", `${file} could not start: ${error.message}`));
    });
    child.once("close", (status) => {
      const result = { status: status ?? 1, stdout, stderr };
      if (result.status !== 0 && !allowFailure) {
        reject(excelLiveError(
          "certificate-command-failed",
          `${file} exited ${result.status}${stderr.trim() ? `: ${stderr.trim()}` : "."}`,
        ));
        return;
      }
      resolve(result);
    });
  });
}
