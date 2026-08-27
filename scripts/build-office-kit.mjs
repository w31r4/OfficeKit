import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const project = path.join(repoRoot, "native", "OfficeKit", "src", "OfficeKit.NativeHost", "OfficeKit.NativeHost.csproj");
const packageMetadata = readJson(path.join(repoRoot, "package.json"));
const targets = Object.freeze({
  "darwin-arm64": Object.freeze({ rid: "osx-arm64", executable: "officekit-codec" }),
  "linux-x64": Object.freeze({ rid: "linux-x64", executable: "officekit-codec" }),
  "win32-x64": Object.freeze({ rid: "win-x64", executable: "officekit-codec.exe" }),
});

const options = parseArguments(process.argv.slice(2));
const target = options.target ?? currentTarget();
const targetConfig = targets[target];
if (!targetConfig) fail(`unsupported target ${target}; expected ${Object.keys(targets).join(", ")}`);
const platformPackageName = `office-kit-codec-${target}`;
const sourcePackageRoot = path.join(repoRoot, "packages", platformPackageName);
const platformPackage = readJson(path.join(sourcePackageRoot, "package.json"));
if (platformPackage.name !== platformPackageName || platformPackage.version !== packageMetadata.version) {
  fail(`${platformPackageName} metadata must match office-kit ${packageMetadata.version}`);
}
const destination = options.output ?? (process.env.OFFICE_KIT_OUTPUT ? path.resolve(process.env.OFFICE_KIT_OUTPUT) : sourcePackageRoot);
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), `office-kit-native-${target}-`));
const publishDirectory = path.join(temporary, "publish");

try {
  run("dotnet", ["clean", project, "--configuration", "Release", "--runtime", targetConfig.rid, "--verbosity", "quiet"]);
  run("dotnet", ["restore", project, "--locked-mode"]);
  run("dotnet", [
    "publish", project,
    "--configuration", "Release",
    "--runtime", targetConfig.rid,
    "--self-contained", "true",
    "--no-restore",
    "--output", publishDirectory,
  ]);

  const stage = path.join(temporary, "package");
  fs.mkdirSync(path.join(stage, "bin"), { recursive: true });
  fs.copyFileSync(path.join(sourcePackageRoot, "package.json"), path.join(stage, "package.json"));
  const executableDestination = path.join(stage, "bin", targetConfig.executable);
  fs.copyFileSync(path.join(publishDirectory, targetConfig.executable), executableDestination);
  if (target !== "win32-x64") fs.chmodSync(executableDestination, 0o755);
  fs.copyFileSync(path.join(repoRoot, "LICENSE"), path.join(stage, "LICENSE"));
  fs.copyFileSync(path.join(repoRoot, "THIRD_PARTY_NOTICES.md"), path.join(stage, "THIRD_PARTY_NOTICES.md"));
  copyDotnetNotices(stage);
  writeSbom(stage, platformPackageName, target, targetConfig.rid);

  const files = listFiles(stage)
    .filter((file) => file !== "manifest.json" && file !== "package.json")
    .map((file) => fileRecord(stage, file));
  const manifest = {
    schemaVersion: 1,
    packageVersion: packageMetadata.version,
    backend: "native-aot",
    transportVersion: 1,
    protocolVersion: 2,
    target,
    runtimeIdentifier: targetConfig.rid,
    targetFramework: "net8.0",
    assemblyName: "officekit-codec",
    sdkVersion: runText("dotnet", ["--version"]),
    sourceProject: "native/OfficeKit/src/OfficeKit.NativeHost/OfficeKit.NativeHost.csproj",
    sourceDependencies: {
      "DocumentFormat.OpenXml": "3.5.1",
      "Google.Protobuf": "3.35.1"
    },
    files,
    totalBytes: files.reduce((sum, file) => sum + file.bytes, 0)
  };
  fs.writeFileSync(path.join(stage, "manifest.json"), `${JSON.stringify(manifest, null, 2)}\n`);
  publishStage(stage, destination);
  console.log(`OfficeKit NativeAOT ${target}: ${files.length} files, ${manifest.totalBytes} bytes`);
} finally {
  fs.rmSync(temporary, { recursive: true, force: true });
}

function parseArguments(args) {
  const parsed = {};
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument !== "--target" && argument !== "--output") fail(`unknown option ${argument}`);
    const value = args[index + 1];
    if (!value || value.startsWith("--")) fail(`${argument} requires a value`);
    const key = argument.slice(2);
    if (parsed[key]) fail(`${argument} may be supplied only once`);
    parsed[key] = argument === "--output" ? path.resolve(value) : value;
    index += 1;
  }
  return parsed;
}

function currentTarget() {
  const target = `${process.platform}-${process.arch}`;
  if (!targets[target]) fail(`current platform ${target} is unsupported`);
  return target;
}

function publishStage(stage, output) {
  fs.mkdirSync(output, { recursive: true });
  const generated = [
    "bin", "manifest.json", "sbom.cdx.json", "LICENSE", "THIRD_PARTY_NOTICES.md",
    "DOTNET-LICENSE.TXT", "DOTNET-THIRD-PARTY-NOTICES.TXT"
  ];
  for (const entry of generated) fs.rmSync(path.join(output, entry), { recursive: true, force: true });
  for (const entry of fs.readdirSync(stage)) {
    if (entry === "package.json" && path.resolve(output) === path.resolve(sourcePackageRoot)) continue;
    fs.cpSync(path.join(stage, entry), path.join(output, entry), { recursive: true });
  }
}

function copyDotnetNotices(stage) {
  const assets = readJson(path.join(repoRoot, "native", "OfficeKit", "src", "OfficeKit.NativeHost", "obj", "project.assets.json"));
  const ilCompiler = Object.keys(assets.libraries || {}).find((name) => name.startsWith("Microsoft.DotNet.ILCompiler/"));
  if (!ilCompiler) fail("NativeAOT compiler package is absent from project assets");
  const [packageName, version] = ilCompiler.split("/");
  const globalPackages = runText("dotnet", ["nuget", "locals", "global-packages", "--list"]).replace(/^global-packages:\s*/u, "").trim();
  const packageRoot = path.join(globalPackages, packageName.toLowerCase(), version);
  fs.copyFileSync(path.join(packageRoot, "LICENSE.TXT"), path.join(stage, "DOTNET-LICENSE.TXT"));
  fs.copyFileSync(path.join(packageRoot, "THIRD-PARTY-NOTICES.TXT"), path.join(stage, "DOTNET-THIRD-PARTY-NOTICES.TXT"));
}

function writeSbom(stage, name, target, rid) {
  const component = (componentName, version, license, purl) => ({
    type: "library",
    name: componentName,
    version,
    scope: "required",
    licenses: [{ license: { id: license } }],
    purl
  });
  const sbom = {
    bomFormat: "CycloneDX",
    specVersion: "1.5",
    version: 1,
    metadata: {
      component: {
        type: "application",
        name,
        version: packageMetadata.version,
        purl: `pkg:npm/${name}@${packageMetadata.version}`,
        properties: [{ name: "office-kit:target", value: target }]
      }
    },
    components: [
      component("Microsoft .NET NativeAOT", "8.0.28", "MIT", `pkg:nuget/Microsoft.DotNet.ILCompiler@8.0.28?rid=${rid}`),
      component("DocumentFormat.OpenXml", "3.5.1", "MIT", "pkg:nuget/DocumentFormat.OpenXml@3.5.1"),
      component("DocumentFormat.OpenXml.Framework", "3.5.1", "MIT", "pkg:nuget/DocumentFormat.OpenXml.Framework@3.5.1"),
      component("Google.Protobuf", "3.35.1", "BSD-3-Clause", "pkg:nuget/Google.Protobuf@3.35.1"),
      component("System.IO.Packaging", "8.0.1", "MIT", "pkg:nuget/System.IO.Packaging@8.0.1")
    ]
  };
  fs.writeFileSync(path.join(stage, "sbom.cdx.json"), `${JSON.stringify(sbom, null, 2)}\n`);
}

function listFiles(root, base = root) {
  return fs.readdirSync(root, { withFileTypes: true }).flatMap((entry) => {
    const target = path.join(root, entry.name);
    return entry.isDirectory() ? listFiles(target, base) : [path.relative(base, target).split(path.sep).join("/")];
  }).sort();
}

function fileRecord(root, file) {
  const bytes = fs.readFileSync(path.join(root, file));
  return { path: file, bytes: bytes.byteLength, sha256: createHash("sha256").update(bytes).digest("hex") };
}

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, "utf8"));
}

function run(command, args) {
  const result = spawnSync(command, args, { cwd: repoRoot, encoding: "utf8", stdio: "inherit", shell: false });
  if (result.status !== 0) process.exit(result.status || 1);
}

function runText(command, args) {
  const result = spawnSync(command, args, { cwd: repoRoot, encoding: "utf8", shell: false });
  if (result.status !== 0) fail(`${command} ${args.join(" ")} failed: ${result.stderr}`);
  return String(result.stdout).trim();
}

function fail(message) {
  throw new Error(`OfficeKit NativeAOT build: ${message}`);
}
