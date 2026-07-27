#!/usr/bin/env node
/**
 * Verify and safely extract one locally-built PDF capability-pack archive.
 *
 * Release workflows use the same strict USTAR extractor as the managed
 * installer. This avoids testing a Windows Git Bash tar implementation in
 * place of the code that customers will actually use.
 */

import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { safeExtractTarGz } from "../src/pdf/providers/installer.mjs";

const PACK_SCHEMA = "office-kit.pdf-provider-pack.v1";
const SHA256 = /^[a-f0-9]{64}$/i;

function fail(message) {
  throw new Error(`PDF capability-pack verification: ${message}`);
}

function plainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function parseArguments(argv) {
  const values = {};
  const allowed = new Set(["archive", "manifest", "destination"]);
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail(`unexpected argument ${token}.`);
    const name = token.slice(2);
    if (!allowed.has(name)) fail(`unknown argument --${name}.`);
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) fail(`--${name} requires a value.`);
    if (Object.hasOwn(values, name)) fail(`--${name} may be supplied only once.`);
    values[name] = value;
    index += 1;
  }
  for (const name of ["archive", "manifest", "destination"]) {
    if (!nonEmptyString(values[name])) fail(`--${name} is required.`);
  }
  return Object.fromEntries(Object.entries(values).map(([name, value]) => [name, path.resolve(value)]));
}

async function regularBytes(filePath, label) {
  const stat = await fs.lstat(filePath);
  if (!stat.isFile() || stat.isSymbolicLink()) fail(`${label} must be a regular non-symlink file: ${filePath}.`);
  return { bytes: await fs.readFile(filePath), size: stat.size };
}

async function emptyRealDirectory(directory) {
  try {
    await fs.mkdir(directory, { recursive: false, mode: 0o700 });
  } catch (error) {
    if (error?.code !== "EEXIST") throw error;
  }
  const stat = await fs.lstat(directory);
  if (!stat.isDirectory() || stat.isSymbolicLink()) fail(`destination must be a real directory: ${directory}.`);
  if ((await fs.readdir(directory)).length) fail(`destination must be empty: ${directory}.`);
}

function parseManifest(bytes, archivePath, archiveSize) {
  let manifest;
  try {
    manifest = JSON.parse(bytes);
  } catch (error) {
    fail(`manifest is not valid JSON: ${error.message}`);
  }
  if (!plainObject(manifest) || manifest.schema !== PACK_SCHEMA || manifest.schemaVersion !== 1) {
    fail("manifest uses an unsupported pack schema.");
  }
  const artifact = manifest.artifact;
  if (!plainObject(artifact) || !nonEmptyString(artifact.asset) || !SHA256.test(artifact.sha256 || "")
    || !Number.isSafeInteger(artifact.downloadBytes) || artifact.downloadBytes <= 0
    || !Number.isSafeInteger(artifact.unpackedBytes) || artifact.unpackedBytes <= 0
    || artifact.archiveFormat !== "tar.gz") {
    fail("manifest has an invalid artifact record.");
  }
  if (artifact.asset !== path.basename(archivePath)) fail("manifest artifact name does not match --archive.");
  if (artifact.downloadBytes !== archiveSize) fail(`archive size is ${archiveSize}; expected ${artifact.downloadBytes}.`);
  if (!plainObject(manifest.payload) || !Array.isArray(manifest.payload.entries)) fail("manifest lacks a payload entry list.");
  if (!plainObject(manifest.sbom) || !SHA256.test(manifest.sbom.sha256 || "") || !Number.isSafeInteger(manifest.sbom.bytes)) fail("manifest has an invalid SBOM record.");
  if (!plainObject(manifest.thirdPartyNotices) || !SHA256.test(manifest.thirdPartyNotices.sha256 || "") || !Number.isSafeInteger(manifest.thirdPartyNotices.bytes)) {
    fail("manifest has an invalid third-party-notices record.");
  }
  return manifest;
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const [archive, manifestFile] = await Promise.all([
    regularBytes(options.archive, "archive"),
    regularBytes(options.manifest, "manifest"),
  ]);
  const manifest = parseManifest(manifestFile.bytes, options.archive, archive.size);
  if (sha256(archive.bytes) !== manifest.artifact.sha256.toLowerCase()) fail("archive SHA-256 does not match its manifest.");

  await emptyRealDirectory(options.destination);
  const extraction = await safeExtractTarGz(archive.bytes, options.destination, manifest.artifact.unpackedBytes);
  const expectedEntries = manifest.payload.entries.map((entry) => entry?.path).sort();
  if (expectedEntries.some((entry) => !nonEmptyString(entry)) || JSON.stringify(extraction.entries) !== JSON.stringify(expectedEntries)) {
    fail("extracted entries do not exactly match the manifest.");
  }
  const [sbom, notices] = await Promise.all([
    regularBytes(path.join(options.destination, "sbom.cdx.json"), "extracted SBOM"),
    regularBytes(path.join(options.destination, "THIRD_PARTY_NOTICES.md"), "extracted third-party notices"),
  ]);
  if (sbom.size !== manifest.sbom.bytes || sha256(sbom.bytes) !== manifest.sbom.sha256.toLowerCase()) fail("extracted SBOM does not match the manifest.");
  if (notices.size !== manifest.thirdPartyNotices.bytes || sha256(notices.bytes) !== manifest.thirdPartyNotices.sha256.toLowerCase()) {
    fail("extracted third-party notices do not match the manifest.");
  }
  process.stdout.write(`${JSON.stringify({
    pack: manifest.pack,
    version: manifest.version,
    platform: manifest.platform,
    artifact: manifest.artifact.asset,
    unpackedBytes: extraction.unpackedBytes,
    entries: extraction.entries.length,
  })}\n`);
}

main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
