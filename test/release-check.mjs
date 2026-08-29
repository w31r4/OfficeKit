import assert from "node:assert/strict";
import fs from "node:fs";
import { spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const result = spawnSync(process.execPath, ["scripts/release-check.mjs", "--json", "--skip-network", "--skip-commands", "--allow-dirty"], {
  cwd: repoRoot,
  encoding: "utf8",
});
assert.equal(result.status, 0, `release-check failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
const report = JSON.parse(result.stdout);
assert.equal(report.package.name, "office-kit");
assert.equal(report.publishReady, true);
assert.ok(report.checks.some((check) => check.name === "package metadata" && check.ok));
assert.ok(report.checks.some((check) => check.name === "project license" && check.ok));
assert.ok(report.checks.some((check) => check.name === "third-party license policy" && check.ok));
assert.ok(report.checks.some((check) =>
  check.name === "standalone release metadata" &&
  check.ok &&
  check.stdout.includes("npm candidate")
));
assert.ok(report.checks.some((check) => check.name === "npm auth" && check.skipped));
assert.match(report.nextPublishCommand, /npm publish/);
const releaseWorkflow = fs.readFileSync(path.join(repoRoot, ".github/workflows/release.yml"), "utf8");
const ciWorkflow = fs.readFileSync(path.join(repoRoot, ".github/workflows/ci.yml"), "utf8");
const slowWorkflow = fs.readFileSync(path.join(repoRoot, ".github/workflows/ci-slow.yml"), "utf8");
const windowsLiveWorkflow = fs.readFileSync(path.join(repoRoot, ".github/workflows/windows-office-live.yml"), "utf8");
const windowsPptxLosslessWorkflow = fs.readFileSync(path.join(repoRoot, ".github/workflows/windows-pptx-lossless.yml"), "utf8");
const standaloneWorkflow = fs.readFileSync(path.join(repoRoot, ".github/workflows/standalone-release.yml"), "utf8");
const dotnetToolchain = JSON.parse(fs.readFileSync(path.join(repoRoot, "global.json"), "utf8"));
assert.equal(dotnetToolchain.sdk.version, "8.0.128");
assert.equal(dotnetToolchain.sdk.rollForward, "disable", "locked OfficeKit restore must not select a newer SDK patch with different implicit build packages");
assert.match(releaseWorkflow, /workflow_dispatch/);
assert.match(releaseWorkflow, /publish_npm/);
assert.match(releaseWorkflow, /default: "false"/);
assert.match(releaseWorkflow, /secrets\.NPM_TOKEN/);
assert.match(releaseWorkflow, /gh release create/);
assert.match(releaseWorkflow, /npm run test:slow/);
assert.match(ciWorkflow, /name:\s*Fast gate/);
assert.match(ciWorkflow, /npm test/);
assert.doesNotMatch(ciWorkflow, /playwright install|libreoffice-writer/);
assert.match(ciWorkflow, /setup-dotnet@v5/);
assert.match(slowWorkflow, /workflow_dispatch/);
assert.match(slowWorkflow, /npm run test:slow/);
const slowSegments = [
  "foundation",
  "presentation",
  "templates",
  "officekit",
  "documents",
  "pdf-packs",
  "pdf-providers",
  "pdf-specialists",
  "qa",
  "release",
];
for (const segment of slowSegments) {
  assert.match(slowWorkflow, new RegExp(`npm run test:slow -- --segment ${segment}`));
  assert.match(releaseWorkflow, new RegExp(`npm run test:slow -- --segment ${segment}`));
}
assert.doesNotMatch(slowWorkflow, /run:\s*npm run test:slow\s*$/m);
assert.doesNotMatch(releaseWorkflow, /run:\s*npm run test:slow\s*$/m);
assert.doesNotMatch(slowWorkflow, /OFFICE_TEMPLATE_SOURCE_ROOT:/);
assert.doesNotMatch(releaseWorkflow, /OFFICE_TEMPLATE_SOURCE_ROOT:/);
assert.match(releaseWorkflow, /submodules:\s*true/);
assert.match(windowsLiveWorkflow, /self-hosted/);
assert.match(windowsLiveWorkflow, /validate-windows-live-evidence/);
assert.match(windowsPptxLosslessWorkflow, /workflow_dispatch/);
assert.match(windowsPptxLosslessWorkflow, /self-hosted/);
assert.match(windowsPptxLosslessWorkflow, /validate-windows-pptx-lossless-evidence/);
assert.match(windowsPptxLosslessWorkflow, /OFFICEKIT_CHECKED_OUT_SHA/);
assert.match(standaloneWorkflow, /node-version:\s*24\.18\.0/);
assert.match(standaloneWorkflow, /darwin-arm64/);
assert.match(standaloneWorkflow, /linux-x64/);
assert.match(standaloneWorkflow, /standalone-four-formats\.mjs/);
assert.match(standaloneWorkflow, /actions\/upload-artifact@v4/);
assert.match(standaloneWorkflow, /gh release upload/);
assert.doesNotMatch(standaloneWorkflow, /push:\s*\n\s*tags:/m);
for (const workflow of [slowWorkflow, releaseWorkflow]) {
  assert.match(workflow, /playwright install --with-deps chromium/);
  assert.match(workflow, /libreoffice-writer libreoffice-calc libreoffice-impress poppler-utils/);
  assert.match(workflow, /actions\/checkout@v5/);
  assert.match(workflow, /actions\/setup-node@v5/);
  assert.match(workflow, /actions\/setup-dotnet@v5/);
  assert.match(workflow, /dotnet-version:\s*8\.0\.128/);
  assert.match(workflow, /soffice --version/);
  assert.match(workflow, /pdfinfo -v/);
  assert.match(workflow, /dotnet test native\/OfficeBridge/);
}

const rejectedPolicyPath = path.join(os.tmpdir(), `office-kit-license-policy-${process.pid}.json`);
const normalPolicy = JSON.parse(fs.readFileSync(path.join(repoRoot, "scripts", "license-policy.json"), "utf8"));
fs.writeFileSync(rejectedPolicyPath, JSON.stringify({ ...normalPolicy, allowedLockLicenses: normalPolicy.allowedLockLicenses.filter((license) => license !== "MIT") }));
try {
  const rejected = spawnSync(process.execPath, ["scripts/release-check.mjs", "--json", "--skip-network", "--skip-commands", "--allow-dirty"], {
    cwd: repoRoot,
    encoding: "utf8",
    env: { ...process.env, OFFICE_KIT_LICENSE_POLICY: rejectedPolicyPath },
  });
  assert.equal(rejected.status, 1);
  const rejectedReport = JSON.parse(rejected.stdout);
  assert.ok(rejectedReport.checks.some((check) => check.name === "third-party license policy" && !check.ok && /unapproved license expression MIT/.test(check.stderr)));
  assert.ok(rejectedReport.blockers.some((blocker) => /license policy/.test(blocker)));
} finally {
  fs.rmSync(rejectedPolicyPath, { force: true });
}

console.log("release check smoke ok");
