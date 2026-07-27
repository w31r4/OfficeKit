import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = new Set(process.argv.slice(2));
const json = args.has("--json");
const skipNetwork = args.has("--skip-network");
const skipCommands = args.has("--skip-commands");
const allowDirty = args.has("--allow-dirty");

function run(command, commandArgs, options = {}) {
  const result = spawnSync(command, commandArgs, {
    cwd: repoRoot,
    encoding: "utf8",
    shell: false,
    ...options,
  });
  return {
    ok: result.status === 0,
    status: result.status,
    stdout: String(result.stdout || "").trim(),
    stderr: String(result.stderr || "").trim(),
    command: [command, ...commandArgs].join(" "),
  };
}

function commandExists(command) {
  const result = spawnSync(process.platform === "win32" ? "where" : "command", process.platform === "win32" ? [command] : ["-v", command], { encoding: "utf8", shell: process.platform !== "win32" });
  return result.status === 0;
}

function summarizeCheck(name, result, required = true) {
  return { name, required, ...result };
}

const pkg = JSON.parse(fs.readFileSync(path.join(repoRoot, "package.json"), "utf8"));
const lock = JSON.parse(fs.readFileSync(path.join(repoRoot, "package-lock.json"), "utf8"));
const licensePolicyPath = process.env.OFFICE_KIT_LICENSE_POLICY || path.join(repoRoot, "scripts", "license-policy.json");
const licensePolicy = JSON.parse(fs.readFileSync(licensePolicyPath, "utf8"));
const checks = [];
const blockers = [];

const gitStatus = run("git", ["status", "--short", "--untracked-files=normal"]);
checks.push(summarizeCheck("git status clean", { ...gitStatus, ok: gitStatus.ok && (allowDirty || !gitStatus.stdout), stdout: gitStatus.stdout || "clean" }, !allowDirty));
if (gitStatus.stdout && !allowDirty) blockers.push("Working tree is not clean.");

const projectLicensePath = path.join(repoRoot, "LICENSE");
const projectLicenseText = fs.existsSync(projectLicensePath) ? fs.readFileSync(projectLicensePath, "utf8") : "";
const projectLicenseOk = pkg.license === "AGPL-3.0-or-later"
  && lock.packages?.[""]?.license === pkg.license
  && pkg.files?.includes("LICENSE")
  && /GNU AFFERO GENERAL PUBLIC LICENSE/.test(projectLicenseText)
  && /Version 3, 19 November 2007/.test(projectLicenseText);
checks.push(summarizeCheck("project license", {
  ok: projectLicenseOk,
  stdout: projectLicenseOk ? pkg.license : "project license audit failed",
  stderr: projectLicenseOk ? "" : "package, lockfile, shipped LICENSE, and canonical GNU AGPL v3 text must agree",
  command: "audit project license metadata + LICENSE",
}));
if (!projectLicenseOk) blockers.push("Project license metadata or GNU AGPL v3 text is incomplete.");

checks.push(summarizeCheck("package metadata", {
  ok: Boolean(pkg.name && pkg.version && pkg.type === "module" && pkg.exports?.["."] && pkg.files?.includes("src/**")),
  stdout: `${pkg.name}@${pkg.version}`,
  stderr: "",
  command: "read package.json",
}));
if (!checks.at(-1).ok) blockers.push("package.json metadata is incomplete for npm publish.");

const standaloneRuntimePath = path.join(repoRoot, "standalone", "node-runtimes.v1.json");
const standaloneReleasePath = path.join(repoRoot, "standalone", "releases.v1.json");
const standaloneInstallerPath = path.join(repoRoot, "standalone", "install.sh");
const standaloneWindowsInstallerPath = path.join(repoRoot, "standalone", "install.ps1");
const standaloneVerifierPath = path.join(repoRoot, "standalone", "verify-install.mjs");
const standaloneWorkflowPath = path.join(repoRoot, ".github", "workflows", "standalone-release.yml");
let standaloneIssues = [];
try {
  const runtimeCatalog = JSON.parse(fs.readFileSync(standaloneRuntimePath, "utf8"));
  const releaseCatalog = JSON.parse(fs.readFileSync(standaloneReleasePath, "utf8"));
  const installer = fs.readFileSync(standaloneInstallerPath, "utf8");
  const windowsInstaller = fs.readFileSync(standaloneWindowsInstallerPath, "utf8");
  const verifier = fs.readFileSync(standaloneVerifierPath, "utf8");
  const workflow = fs.readFileSync(standaloneWorkflowPath, "utf8");
  if (runtimeCatalog.schemaVersion !== 1 || runtimeCatalog.nodeVersion !== "24.18.0") {
    standaloneIssues.push("Node runtime catalog must pin schema 1 and Node 24.18.0");
  }
  if (releaseCatalog.schemaVersion !== 1 || releaseCatalog.officeKitVersion !== pkg.version) {
    standaloneIssues.push("standalone release catalog version does not match package.json");
  }
  for (const target of ["darwin-arm64", "linux-x64", "win32-x64"]) {
    const runtime = runtimeCatalog.runtimes?.[target];
    const release = releaseCatalog.assets?.[target];
    const archiveExtension = target === "win32-x64" ? ".zip" : ".tar.gz";
    const targetInstaller = target === "win32-x64" ? windowsInstaller : installer;
    if (
      !runtime ||
      !/^https:\/\/nodejs\.org\/dist\/v24\.18\.0\//.test(runtime.url) ||
      !/^[a-f0-9]{64}$/.test(runtime.sha256) ||
      !Number.isSafeInteger(runtime.size) ||
      runtime.size <= 0
    ) {
      standaloneIssues.push(`${target} Node runtime pin is incomplete`);
    }
    if (
      !release ||
      release.asset !== `office-kit-${pkg.version}-${target}${archiveExtension}` ||
      !/^[a-f0-9]{64}$/.test(release.sha256) ||
      !Number.isSafeInteger(release.size) ||
      release.size <= 0
    ) {
      standaloneIssues.push(`${target} standalone release pin is incomplete`);
    } else if (
      !targetInstaller.includes(release.sha256) ||
      !targetInstaller.includes(String(release.size))
    ) {
      standaloneIssues.push(`${target} installer constants do not match the release catalog`);
    }
  }
  if (
    !installer.includes(`OFFICE_KIT_VERSION=${pkg.version}`) ||
    /FINALIZE_/.test(installer) ||
    (fs.statSync(standaloneInstallerPath).mode & 0o111) === 0
  ) {
    standaloneIssues.push("standalone installer version, hashes, or executable mode are incomplete");
  }
  if (
    !windowsInstaller.includes(`$OfficeKitVersion = "${pkg.version}"`) ||
    /RELEASE_(?:SHA256|SIZE)/.test(windowsInstaller) ||
    !/Invoke-WebRequest/.test(windowsInstaller) ||
    !/Get-FileHash/.test(windowsInstaller)
  ) {
    standaloneIssues.push("Windows standalone installer version, hashes, or verification is incomplete");
  }
  if (
    !/standalone-manifest\.json/.test(verifier) ||
    !/installed file failed integrity verification/.test(verifier) ||
    (fs.statSync(standaloneVerifierPath).mode & 0o111) === 0
  ) {
    standaloneIssues.push("standalone installed-file verifier is incomplete");
  }
  if (
    !/linux-x64/.test(workflow) ||
    !/darwin-arm64/.test(workflow) ||
    !/node-version:\s*24\.18\.0/.test(workflow) ||
    !/win32-x64/.test(workflow) ||
    !/install\.ps1/.test(workflow) ||
    !/standalone-four-formats\.mjs/.test(workflow)
  ) {
    standaloneIssues.push("standalone release workflow does not verify all native targets and all four formats");
  }
} catch (error) {
  standaloneIssues.push(error instanceof Error ? error.message : String(error));
}
const standaloneOk = standaloneIssues.length === 0;
checks.push(summarizeCheck("standalone release metadata", {
  ok: standaloneOk,
  stdout: standaloneOk ? `OfficeKit ${pkg.version}, Node 24.18.0, darwin-arm64 + linux-x64 + win32-x64` : "standalone release audit failed",
  stderr: standaloneIssues.join("\n"),
  command: "audit standalone runtime, release, installer, and workflow pins",
}));
if (!standaloneOk) blockers.push("Self-contained release metadata is incomplete or inconsistent.");

const noticesPath = path.join(repoRoot, "THIRD_PARTY_NOTICES.md");
const declaredDependencyNames = [...new Set([...Object.keys(pkg.dependencies || {}), ...Object.keys(pkg.peerDependencies || {})])];
const policyNames = Object.keys(licensePolicy.declaredPackages || {});
const lockLicenseIssues = Object.entries(lock.packages || {}).flatMap(([packagePath, metadata]) => {
  if (!packagePath.startsWith("node_modules/")) return [];
  if (!metadata.license) return [`${packagePath}: missing license metadata`];
  if (!licensePolicy.allowedLockLicenses.includes(metadata.license)) return [`${packagePath}: unapproved license expression ${metadata.license}`];
  return [];
});
const missingPolicy = declaredDependencyNames.filter((name) => !policyNames.includes(name));
const stalePolicy = policyNames.filter((name) => !declaredDependencyNames.includes(name));
const noticesText = fs.existsSync(noticesPath) ? fs.readFileSync(noticesPath, "utf8") : "";
const missingNotices = declaredDependencyNames.filter((name) => !noticesText.toLowerCase().includes(name.toLowerCase()));
const licenseOk = fs.existsSync(noticesPath) && pkg.files?.includes("THIRD_PARTY_NOTICES.md") && !lockLicenseIssues.length && !missingPolicy.length && !stalePolicy.length && !missingNotices.length;
checks.push(summarizeCheck("third-party license policy", {
  ok: licenseOk,
  stdout: licenseOk ? `${Object.keys(lock.packages || {}).filter((name) => name.startsWith("node_modules/")).length} locked packages audited` : "license audit failed",
  stderr: [...lockLicenseIssues, ...missingPolicy.map((name) => `missing policy: ${name}`), ...stalePolicy.map((name) => `stale policy: ${name}`), ...missingNotices.map((name) => `missing notice: ${name}`)].join("\n"),
  command: "audit package-lock.json + THIRD_PARTY_NOTICES.md",
}));
if (!licenseOk) blockers.push("Third-party license policy or notices are incomplete.");

if (!skipCommands) {
  for (const [name, commandArgs] of [
    ["npm test", ["test"]],
    ["npm run docs:api", ["run", "docs:api"]],
    ["npm run test:pack", ["run", "test:pack"]],
  ]) {
    const check = summarizeCheck(name, run("npm", commandArgs));
    checks.push(check);
    if (!check.ok) blockers.push(`${name} failed.`);
  }
  if (fs.existsSync(path.join(repoRoot, "native", "OfficeBridge")) && commandExists("dotnet")) {
    const check = summarizeCheck("dotnet test native/OfficeBridge", run("dotnet", ["test", "native/OfficeBridge"]));
    checks.push(check);
    if (!check.ok) blockers.push("dotnet test native/OfficeBridge failed.");
  } else {
    checks.push(summarizeCheck("dotnet test native/OfficeBridge", { ok: true, stdout: "skipped: dotnet or native/OfficeBridge unavailable", stderr: "", command: "dotnet test native/OfficeBridge" }, false));
  }
  if (fs.existsSync(path.join(repoRoot, "native", "OfficeKit")) && commandExists("dotnet")) {
    const check = summarizeCheck("dotnet test native/OfficeKit", run("dotnet", ["test", "native/OfficeKit/OfficeKit.sln", "--configuration", "Release"]));
    checks.push(check);
    if (!check.ok) blockers.push("dotnet test native/OfficeKit failed.");
  } else {
    checks.push(summarizeCheck("dotnet test native/OfficeKit", { ok: true, stdout: "skipped: dotnet or native/OfficeKit unavailable", stderr: "", command: "dotnet test native/OfficeKit" }, false));
  }
}

let npmAuth = { ok: false, skipped: skipNetwork, stdout: "", stderr: "", command: "npm whoami" };
let npmView = { ok: false, skipped: skipNetwork, stdout: "", stderr: "", command: `npm view ${pkg.name} version --json` };
if (!skipNetwork) {
  npmAuth = run("npm", ["whoami"]);
  npmView = run("npm", ["view", pkg.name, "version", "--json"]);
  checks.push(summarizeCheck("npm auth", npmAuth));
  checks.push(summarizeCheck("npm package lookup", npmView, false));
  if (!npmAuth.ok) blockers.push("npm auth unavailable: run npm adduser or configure an npm token before publishing.");
} else {
  checks.push(summarizeCheck("npm auth", npmAuth, false));
  checks.push(summarizeCheck("npm package lookup", npmView, false));
}

const publishedVersion = npmView.ok ? npmView.stdout.replace(/^"|"$/g, "") : null;
if (publishedVersion === pkg.version) blockers.push(`npm ${pkg.name}@${pkg.version} is already published; bump version before publishing.`);

const result = {
  package: { name: pkg.name, version: pkg.version },
  publishReady: blockers.length === 0,
  npmAuth: npmAuth.ok ? npmAuth.stdout : null,
  publishedVersion,
  checks,
  blockers,
  nextPublishCommand: `npm publish --access public`,
};

if (json) {
  console.log(JSON.stringify(result, null, 2));
} else {
  console.log(`${pkg.name}@${pkg.version} release check`);
  for (const check of checks) console.log(`${check.ok ? "✓" : check.required ? "✗" : "-"} ${check.name}${check.stdout ? ` — ${check.stdout.split("\n")[0]}` : ""}`);
  if (blockers.length) {
    console.log("\nBlockers:");
    for (const blocker of blockers) console.log(`- ${blocker}`);
  } else {
    console.log(`\nPublish-ready. Command: ${result.nextPublishCommand}`);
  }
}

process.exit(result.publishReady ? 0 : 1);
