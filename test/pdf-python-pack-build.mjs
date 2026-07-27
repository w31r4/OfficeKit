import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

const root = path.resolve(import.meta.dirname, "..");
const buildScript = path.join(root, "scripts", "build-python-provider-pack.mjs");
const inputPath = path.join(root, "scripts", "pdf-provider-python-release-inputs.v1.json");
const workflowPath = path.join(root, ".github", "workflows", "pdf-python-capability-packs.yml");
const buildSource = await fs.readFile(buildScript, "utf8");
const workflow = await fs.readFile(workflowPath, "utf8");

assert.match(buildSource, /_tkinter/);
assert.match(buildSource, /lib-dynload/);
assert.match(buildSource, /python\.exe/);
assert.match(buildSource, /win32-x64/);
assert.match(buildSource, /Windows runtime extraction requires SystemRoot/);
assert.match(buildSource, /path\.join\(systemRoot, "System32", "tar\.exe"\)/);
assert.match(buildSource, /await extractRuntimeArchive\(runtimeArchive, runtimeExtract, options\.platform\)/);
assert.match(workflow, /platform: win32-x64/);
assert.match(workflow, /runner: windows-2025/);
assert.match(workflow, /PYTHON_PACK_VERSION: 3\.13\.14-oat\.2/);
assert.match(workflow, /--expected-platforms darwin-arm64,linux-x64,win32-x64/);
assert.match(workflow, /python="\$destination\/python\.exe"/);
assert.match(workflow, /verify-pdf-provider-pack\.mjs/);
assert.doesNotMatch(workflow, /tar -xzf/);

function run(arguments_, { expect = 0 } = {}) {
  const result = spawnSync(process.execPath, [buildScript, ...arguments_], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, expect, result.stderr || result.stdout);
  return result;
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-python-pack-build-"));
try {
  const bytes = await fs.readFile(inputPath);
  const source = JSON.parse(bytes);
  const verified = JSON.parse(run(["--verify-lock"]).stdout);
  assert.equal(verified.schema, source.schema);
  assert.equal(verified.sha256, crypto.createHash("sha256").update(bytes).digest("hex"));
  assert.deepEqual(verified.packs, {
    "python-foundation": { "darwin-arm64": 10, "linux-x64": 10, "win32-x64": 10 },
    "python-specialists": { "darwin-arm64": 20, "linux-x64": 20, "win32-x64": 20 },
    "ocr-core": { "darwin-arm64": 27, "linux-x64": 27, "win32-x64": 27 },
  });
  assert.equal(source.pythonRuntime.platforms["win32-x64"].sha256, "91ea0cd883295458fa766eae36241df3ecc21f7029814189707d54b450e55c69");

  async function rejectMutation(label, mutate, expected) {
    const candidate = structuredClone(source);
    mutate(candidate);
    const candidatePath = path.join(temporary, `${label}.json`);
    await fs.writeFile(candidatePath, `${JSON.stringify(candidate, null, 2)}\n`, "utf8");
    const rejected = run(["--verify-lock", "--inputs", candidatePath], { expect: 2 });
    assert.match(rejected.stderr, expected);
  }

  await rejectMutation("duplicate-platform-wheel", (candidate) => {
    candidate.packs["python-foundation"].platformWheels["darwin-arm64"].push(
      structuredClone(candidate.packs["python-foundation"].commonWheels[0]),
    );
  }, /duplicate darwin-arm64 wheel reportlab@4\.4\.9/);
  await rejectMutation("missing-direct-requirement", (candidate) => {
    candidate.packs["python-foundation"].directRequirements.pypdf = "0.0.0";
  }, /direct requirement pypdf==0\.0\.0 is absent from its darwin-arm64 wheel lock/);
  await rejectMutation("unsafe-source-url", (candidate) => {
    candidate.pythonRuntime.platforms["darwin-arm64"].url = "http://example.test/python.tar.gz";
  }, /credential-free HTTPS URL/);
  await rejectMutation("unsupported-wheel-platform", (candidate) => {
    candidate.packs["python-foundation"].platformWheels["unsupported-x64"] = [];
  }, /unsupported platform/);
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("python PDF provider pack build smoke ok");
