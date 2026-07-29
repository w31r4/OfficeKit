#!/usr/bin/env node

/**
 * Sign one unsigned source PDF with the deliberately bounded PromptBench
 * PAdES-LTA test profile. This is not a generic TSA client or a production
 * PAdES workflow: every credential, trust root, CRL, placement, profile flag,
 * and postflight requirement is fixed by the explicit caller input.
 */

import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const MAX_PDF_BYTES = 512 * 1024 * 1024;
const MAX_CREDENTIAL_BYTES = 16 * 1024 * 1024;
const MAX_EVIDENCE_BYTES = 4 * 1024 * 1024;
const fixedTestTsaMoment = "2030-01-01T00:00:00+00:00";
const scriptDirectory = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const providerScripts = path.join(scriptDirectory, "scripts");

function fail(message) {
  throw new Error(message);
}

function usage() {
  return [
    "Usage:",
    "  node officekit-pades-ltv-test-sign-workflow.mjs <source.pdf> <signed.pdf> <audit.json> --signer <signer.p12> --tsa <test-tsa.p12> --root <test-root.pem> --crl <test-root.crl> [--python <provider-python>]",
    "",
    "This is the public, test-only offline PAdES-LTA profile. It never sends a credential to a network TSA and never claims PAdES conformance.",
  ].join("\n");
}

function parseArgs(argv) {
  const [source, output, audit, ...rest] = argv;
  if (!source || !output || !audit) fail(usage());
  const options = {};
  for (let index = 0; index < rest.length; index += 1) {
    const flag = rest[index];
    if (!flag.startsWith("--")) fail(`unexpected positional argument: ${flag}\n${usage()}`);
    const name = flag.slice(2);
    if (!new Set(["signer", "tsa", "root", "crl", "python"]).has(name)) fail(`unsupported option: ${flag}\n${usage()}`);
    const value = rest[++index];
    if (!value || value.startsWith("--")) fail(`${flag} requires one value\n${usage()}`);
    options[name] = value;
  }
  for (const name of ["signer", "tsa", "root", "crl"]) if (!options[name]) fail(`--${name} is required\n${usage()}`);
  return { source, output, audit, ...options };
}

async function regularFile(value, label, maximum, minimum = 1) {
  const candidate = path.resolve(value);
  const stat = await fs.lstat(candidate).catch(() => fail(`${label} does not exist: ${candidate}`));
  if (stat.isSymbolicLink()) fail(`${label} is a symbolic link and will not be followed: ${candidate}`);
  if (!stat.isFile()) fail(`${label} must be a regular file: ${candidate}`);
  if (stat.size < minimum || stat.size > maximum) fail(`${label} size ${stat.size} is outside ${minimum}..${maximum} bytes`);
  return candidate;
}

async function absentOutput(value, label, source) {
  const candidate = path.resolve(value);
  if (candidate === source) fail(`${label} must be distinct from the source PDF`);
  const parent = path.dirname(candidate);
  const parentStat = await fs.lstat(parent).catch(() => fail(`${label} parent does not exist: ${parent}`));
  if (!parentStat.isDirectory() || parentStat.isSymbolicLink()) fail(`${label} parent must be an ordinary directory: ${parent}`);
  try {
    const stat = await fs.lstat(candidate);
    if (stat.isSymbolicLink()) fail(`${label} path is a symbolic link and will not be followed: ${candidate}`);
    fail(`${label} already exists and will not be replaced: ${candidate}`);
  } catch (error) {
    if (error?.code === "ENOENT") return candidate;
    throw error;
  }
}

async function sha256(file) {
  const bytes = await fs.readFile(file);
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function readPrefix(file, length) {
  const handle = await fs.open(file, "r");
  try {
    const target = Buffer.alloc(length);
    const { bytesRead } = await handle.read(target, 0, length, 0);
    return target.subarray(0, bytesRead);
  } finally {
    await handle.close();
  }
}

async function exactPrefixPreserved(source, output) {
  const sourceInfo = await fs.stat(source);
  const outputInfo = await fs.stat(output);
  if (outputInfo.size < sourceInfo.size) return false;
  const sourceHandle = await fs.open(source, "r");
  const outputHandle = await fs.open(output, "r");
  try {
    const chunkSize = 1024 * 1024;
    const left = Buffer.alloc(chunkSize);
    const right = Buffer.alloc(chunkSize);
    for (let offset = 0; offset < sourceInfo.size; offset += chunkSize) {
      const size = Math.min(chunkSize, sourceInfo.size - offset);
      const [{ bytesRead: leftRead }, { bytesRead: rightRead }] = await Promise.all([
        sourceHandle.read(left, 0, size, offset),
        outputHandle.read(right, 0, size, offset),
      ]);
      if (leftRead !== size || rightRead !== size || !left.subarray(0, size).equals(right.subarray(0, size))) return false;
    }
    return true;
  } finally {
    await Promise.all([sourceHandle.close(), outputHandle.close()]);
  }
}

function run(command, args, label) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1" },
    maxBuffer: 16 * 1024 * 1024,
  });
  if (result.error) fail(`${label} could not start: ${result.error.message}`);
  if (result.status !== 0) fail(`${label} failed (${result.status}): ${(result.stderr || result.stdout || "unknown failure").trim()}`);
  const text = String(result.stdout || "").trim();
  try { return text ? JSON.parse(text) : {}; } catch (error) { fail(`${label} did not return JSON: ${error.message}`); }
}

async function writeAuditAtomically(target, value) {
  const temporary = path.join(path.dirname(target), `.${path.basename(target)}.${process.pid}.${crypto.randomUUID()}.tmp`);
  await fs.writeFile(temporary, `${JSON.stringify(value, null, 2)}\n`, { mode: 0o600 });
  await fs.rename(temporary, target);
}

async function removeIfPresent(target) {
  await fs.unlink(target).catch((error) => {
    if (error.code !== "ENOENT") throw error;
  });
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const source = await regularFile(options.source, "source PDF", MAX_PDF_BYTES, 5);
  const output = await absentOutput(options.output, "signed PDF output", source);
  const audit = await absentOutput(options.audit, "audit output", source);
  if (audit === output) fail("audit output must be distinct from signed PDF output");
  const signer = await regularFile(options.signer, "signer PKCS#12", MAX_CREDENTIAL_BYTES);
  const tsa = await regularFile(options.tsa, "test TSA PKCS#12", MAX_CREDENTIAL_BYTES);
  const root = await regularFile(options.root, "test trust root", MAX_EVIDENCE_BYTES);
  const crl = await regularFile(options.crl, "test CRL", MAX_EVIDENCE_BYTES);
  const sourceHead = await readPrefix(source, 5);
  if (!sourceHead.equals(Buffer.from("%PDF-"))) fail("source PDF does not begin with a PDF header");
  const python = options.python || process.env.OFFICE_KIT_PDF_PROVIDER_PYTHON;
  if (!python) fail("select the explicitly managed pyHanko runtime with --python or OFFICE_KIT_PDF_PROVIDER_PYTHON; no arbitrary Python fallback is used");

  const hashes = Object.fromEntries(await Promise.all([
    ["source", source], ["signer", signer], ["tsa", tsa], ["root", root], ["crl", crl],
  ].map(async ([name, file]) => [name, await sha256(file)])));
  const signProvider = path.join(providerScripts, "pyhanko_sign_provider.py");
  const verifyProvider = path.join(providerScripts, "pyhanko_provider.py");
  const probe = run(python, [signProvider, "probe"], "pyHanko signing capability probe");
  if (probe?.provider !== "pyhanko" || probe?.silentFallback !== false || !Array.isArray(probe?.ltvEmbeddingSupported) || !probe.ltvEmbeddingSupported.includes("local-test-pades-lta-only")) {
    fail("selected provider does not expose the bounded local-test PAdES-LTA capability");
  }

  let signed;
  let verified;
  let renderPages = [];
  try {
    signed = run(python, [
      signProvider, "sign", source, output,
      "--expected-sha256", hashes.source,
      "--credential", signer, "--credential-sha256", hashes.signer, "--no-passphrase",
      "--field-name", "ApprovalSignature", "--field-mode", "create-visible", "--page-index", "1", "--box", "72,72,300,150",
      "--signature-kind", "approval", "--subfilter", "pades", "--expected-signature-count", "0",
      "--pades-ltv-test-profile",
      "--test-tsa-credential", tsa, "--test-tsa-credential-sha256", hashes.tsa, "--test-tsa-no-passphrase",
      "--ltv-trust-root", root, "--ltv-trust-root-sha256", hashes.root,
      "--ltv-crl", crl, "--ltv-crl-sha256", hashes.crl,
      "--test-tsa-moment", fixedTestTsaMoment, "--caller-isolated",
    ], "bounded PAdES-LTA test signing");
    const outputHash = await sha256(output);
    verified = run(python, [
      verifyProvider, "verify", output, "--expected-sha256", outputHash,
      "--trust-policy", "explicit-roots", "--trust-root", root, "--crl", crl,
      "--revocation-policy", "require", "--require-signature", "--require-all-integrity-valid", "--require-all-trusted",
      "--require-all-bottom-line", "--require-signature-timestamp", "--require-document-timestamp",
      "--require-dss-validation-info", "--require-revocation-evidence",
    ], "bounded PAdES-LTA postflight verification");
    if (verified?.ok !== true || verified?.summary?.signatureCount !== 2 || verified?.summary?.hasDocumentTimestamp !== true || verified?.summary?.allSignatureTimestampsValid !== true) {
      fail("postflight validation did not establish the bounded local test profile");
    }
    const renderRoot = await fs.mkdtemp(path.join(os.tmpdir(), "officekit-pades-ltv-render-"));
    try {
      const render = spawnSync(process.env.OFFICE_KIT_AGENT_EVAL_PDFTOPPM || "pdftoppm", ["-png", "-r", "144", output, path.join(renderRoot, "page")], { encoding: "utf8", maxBuffer: 4 * 1024 * 1024 });
      if (render.error || render.status !== 0) fail(`Poppler render failed: ${(render.stderr || render.error?.message || "unknown failure").trim()}`);
      renderPages = (await fs.readdir(renderRoot)).filter((entry) => /^page-\d+\.png$/.test(entry)).sort();
      if (renderPages.length !== 2) fail(`Poppler render produced ${renderPages.length} pages, expected 2`);
      for (const name of renderPages) {
        const info = await fs.stat(path.join(renderRoot, name));
        if (info.size < 1_000) fail(`Poppler render is unexpectedly small: ${name}`);
      }
    } finally {
      await fs.rm(renderRoot, { recursive: true, force: true });
    }
    if (await sha256(output) !== outputHash) fail("signed PDF changed after postflight validation");
    if (!await exactPrefixPreserved(source, output)) fail("signed PDF does not preserve the exact source revision prefix");
    const sourceBytes = await fs.stat(source);
    const outputBytes = await fs.stat(output);
    const report = {
      schema: "office-kit.pades-ltv-test-sign-workflow.v1",
      ok: true,
      operationCompleted: true,
      status: "success",
      operation: "sign-local-pkcs12-pades-ltv-test-profile",
      provider: { name: "pyhanko", version: probe.providerVersion, certvalidatorVersion: probe.certvalidatorVersion },
      savePolicy: { strategy: "incremental", sourcePrefixPreserved: true },
      silentFallback: false,
      networkAllowed: false,
      source: { path: options.source, sha256: hashes.source, bytes: sourceBytes.size },
      output: { path: options.output, sha256: outputHash, bytes: outputBytes.size },
      credential: { path: options.signer, sha256: hashes.signer, passphraseChannel: "none", secretLogged: false, testOnly: true },
      padesLtvTestProfile: {
        enabled: true,
        testOnly: true,
        networkAllowed: false,
        padesProfileConformanceClaimed: false,
        testTsaMoment: fixedTestTsaMoment,
        tsa: { path: options.tsa, sha256: hashes.tsa },
        trustRoot: { path: options.root, sha256: hashes.root },
        crl: { path: options.crl, sha256: hashes.crl },
      },
      transaction: { noReplace: true, outputPublishedAtomically: true, sourceImmutable: true },
      validation: {
        postflight: { ok: verified.ok === true, summary: verified.summary, dss: verified.dss, ltvEvidence: verified.ltvEvidence },
        poppler: { ok: true, renderer: "pdftoppm", pageCount: renderPages.length },
      },
      limitations: [
        "This is a disclosed, disposable, offline PromptBench test profile.",
        "It does not route to an external TSA or claim complete PAdES profile conformance.",
      ],
    };
    await writeAuditAtomically(audit, report);
    process.stdout.write(`${JSON.stringify({ ok: true, output: report.output, audit: { path: options.audit, sha256: await sha256(audit) } })}\n`);
  } catch (error) {
    await removeIfPresent(output);
    await removeIfPresent(audit);
    throw error;
  }
}

main().catch((error) => {
  process.stderr.write(`officekit-pades-ltv-test-sign-workflow: ${error.message}\n`);
  process.exitCode = 2;
});
