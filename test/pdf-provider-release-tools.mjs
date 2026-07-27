import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

const root = path.resolve(import.meta.dirname, "..");
const hashScript = path.join(root, "scripts", "sha256-file.mjs");
const packBuilder = path.join(root, "scripts", "build-pdf-provider-pack.mjs");
const verifier = path.join(root, "scripts", "verify-pdf-provider-pack.mjs");
const qpdfWorkflow = await fs.readFile(path.join(root, ".github", "workflows", "pdf-capability-packs.yml"), "utf8");

assert.match(qpdfWorkflow, /sha256-file\.mjs/);
assert.match(qpdfWorkflow, /verify-pdf-provider-pack\.mjs/);
assert.doesNotMatch(qpdfWorkflow, /shasum/, "the Windows qpdf lane must not depend on Git Bash's optional Perl shim");

function run(script, arguments_, { expect = 0 } = {}) {
  const result = spawnSync(process.execPath, [script, ...arguments_], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, expect, result.stderr || result.stdout);
  return result;
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-pdf-release-tools-"));
try {
  const input = path.join(temporary, "input.txt");
  await fs.writeFile(input, "hash-pinned release input\n", "utf8");
  const expectedHash = crypto.createHash("sha256").update(await fs.readFile(input)).digest("hex");
  assert.equal(run(hashScript, [input]).stdout.trim(), expectedHash);
  assert.match(run(hashScript, [], { expect: 2 }).stderr, /usage is sha256-file\.mjs/);

  const payload = path.join(temporary, "payload");
  const output = path.join(temporary, "output");
  const destination = path.join(temporary, "consumer");
  const notices = path.join(temporary, "notices.md");
  await fs.mkdir(path.join(payload, "bin"), { recursive: true });
  await fs.writeFile(path.join(payload, "bin", "tool"), "fixture tool\n", { mode: 0o755 });
  await fs.writeFile(notices, "fixture notices\n", "utf8");
  const manifest = JSON.parse(run(packBuilder, [
    "--pack", "fixture-pack",
    "--version", "1.2.3",
    "--platform", "win32-x64",
    "--payload", payload,
    "--output", output,
    "--source-url", "https://example.test/fixture.tar.gz",
    "--source-sha256", "a".repeat(64),
    "--license", "Apache-2.0",
    "--notices", notices,
  ]).stdout);
  const archive = path.join(output, manifest.artifact.asset);
  const manifestPath = path.join(output, "fixture-pack-1.2.3-win32-x64.manifest.json");
  const verified = JSON.parse(run(verifier, ["--archive", archive, "--manifest", manifestPath, "--destination", destination]).stdout);
  assert.deepEqual(verified, {
    pack: "fixture-pack",
    version: "1.2.3",
    platform: "win32-x64",
    artifact: manifest.artifact.asset,
    unpackedBytes: manifest.artifact.unpackedBytes,
    entries: manifest.payload.entries.length,
  });
  assert.equal(await fs.readFile(path.join(destination, "bin", "tool"), "utf8"), "fixture tool\n");

  const tamperedDirectory = path.join(temporary, "tampered-output");
  await fs.mkdir(tamperedDirectory);
  const tampered = path.join(tamperedDirectory, manifest.artifact.asset);
  const bytes = await fs.readFile(archive);
  bytes[0] ^= 0xff;
  await fs.writeFile(tampered, bytes);
  assert.match(
    run(verifier, ["--archive", tampered, "--manifest", manifestPath, "--destination", path.join(temporary, "tampered")], { expect: 2 }).stderr,
    /archive SHA-256 does not match its manifest/,
  );
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("PDF capability-pack release tools smoke ok");
