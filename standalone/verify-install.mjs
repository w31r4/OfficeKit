#!/usr/bin/env node

import crypto from "node:crypto";
import { lstat, readFile, readdir, realpath } from "node:fs/promises";
import path from "node:path";

const RECEIPT = ".office-kit-install-receipt";

function fail(message) {
  throw new Error(`OfficeKit installation verification: ${message}`);
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
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
    normalized === value &&
    normalized !== "." &&
    normalized !== ".." &&
    !normalized.startsWith("../") &&
    !normalized.includes("/../")
  );
}

async function main() {
  const [requestedRoot, expectedVersion, expectedTarget, receiptFlag] =
    process.argv.slice(2);
  if (!requestedRoot || !expectedVersion || !expectedTarget) {
    fail("usage: verify-install.mjs <root> <version> <target> [--allow-receipt].");
  }
  if (receiptFlag != null && receiptFlag !== "--allow-receipt") {
    fail(`unknown option ${receiptFlag}.`);
  }
  const root = await realpath(requestedRoot);
  const manifestPath = path.join(root, "standalone-manifest.json");
  const manifestStat = await lstat(manifestPath);
  if (!manifestStat.isFile() || manifestStat.isSymbolicLink()) {
    fail("standalone-manifest.json must be a regular non-symlink file.");
  }
  const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  if (
    manifest.schema !== "office-kit.standalone.v1" ||
    manifest.officeKitVersion !== expectedVersion ||
    manifest.target !== expectedTarget ||
    !Array.isArray(manifest.files) ||
    manifest.fileCount !== manifest.files.length
  ) {
    fail("manifest identity or file count does not match this installation.");
  }

  const expected = new Map();
  let expectedBytes = 0;
  for (const record of manifest.files) {
    if (
      !safeRelativePath(record?.path) ||
      !Number.isSafeInteger(record.bytes) ||
      record.bytes < 0 ||
      !/^[a-f0-9]{64}$/u.test(record.sha256) ||
      !["0644", "0755"].includes(record.mode) ||
      expected.has(record.path)
    ) {
      fail(`manifest contains an invalid file record: ${JSON.stringify(record)}.`);
    }
    expected.set(record.path, record);
    expectedBytes += record.bytes;
  }
  if (manifest.unpackedBytes !== expectedBytes) {
    fail("manifest unpacked-byte total is inconsistent.");
  }

  const seen = new Set();
  async function walk(directory, relative = "") {
    const children = await readdir(directory, { withFileTypes: true });
    children.sort((left, right) =>
      left.name < right.name ? -1 : left.name > right.name ? 1 : 0,
    );
    for (const child of children) {
      const childRelative = relative ? `${relative}/${child.name}` : child.name;
      if (!safeRelativePath(childRelative)) {
        fail(`installation contains unsafe path ${childRelative}.`);
      }
      const absolute = path.join(directory, child.name);
      const metadata = await lstat(absolute);
      if (metadata.isSymbolicLink()) {
        fail(`installation contains a symlink: ${childRelative}.`);
      }
      if (metadata.isDirectory()) {
        await walk(absolute, childRelative);
        continue;
      }
      if (!metadata.isFile()) {
        fail(`installation contains a special file: ${childRelative}.`);
      }
      if (childRelative === "standalone-manifest.json") continue;
      if (
        childRelative === RECEIPT &&
        receiptFlag === "--allow-receipt"
      ) {
        continue;
      }
      const record = expected.get(childRelative);
      if (!record) fail(`installation contains untracked file ${childRelative}.`);
      const bytes = await readFile(absolute);
      const mode = (metadata.mode & 0o111) !== 0 ? "0755" : "0644";
      if (
        bytes.length !== record.bytes ||
        sha256(bytes) !== record.sha256 ||
        mode !== record.mode
      ) {
        fail(`installed file failed integrity verification: ${childRelative}.`);
      }
      seen.add(childRelative);
    }
  }
  await walk(root);
  for (const pathname of expected.keys()) {
    if (!seen.has(pathname)) fail(`installed file is missing: ${pathname}.`);
  }
  process.stdout.write(
    `${JSON.stringify({
      ok: true,
      version: expectedVersion,
      target: expectedTarget,
      files: seen.size,
      bytes: expectedBytes,
    })}\n`,
  );
}

main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = 2;
});
