import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  buildStandalone,
  createDeterministicTarGz,
  createDeterministicZip,
} from "../scripts/build-standalone.mjs";

const repositoryRoot = path.resolve(import.meta.dirname, "..");
const packageMetadata = JSON.parse(
  await fs.readFile(path.join(repositoryRoot, "package.json"), "utf8"),
);

const target =
  process.platform === "darwin" && process.arch === "arm64"
    ? "darwin-arm64"
    : process.platform === "linux" && process.arch === "x64"
      ? "linux-x64"
      : process.platform === "win32" && process.arch === "x64"
        ? "win32-x64"
      : null;
const windows = target === "win32-x64";
const anydocNativePackage = {
  "darwin-arm64": "@firecrawl/anydoc-darwin-arm64",
  "linux-x64": "@firecrawl/anydoc-linux-x64-gnu",
  "win32-x64": "@firecrawl/anydoc-win32-x64-msvc",
}[target];
const installer = path.join(
  repositoryRoot,
  "standalone",
  windows ? "install.ps1" : "install.sh",
);

if (target == null) {
  console.log(`standalone distribution smoke skipped on ${process.platform}-${process.arch}`);
  process.exit(0);
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function execute(command, args, { cwd = repositoryRoot, env = {}, expect = 0 } = {}) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    env: { ...process.env, ...env },
    maxBuffer: 64 * 1024 * 1024,
    shell: process.platform === "win32" && /\.(?:cmd|bat)$/iu.test(command),
  });
  assert.equal(
    result.status,
    expect,
    `${command} ${args.join(" ")} exited ${result.status}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`,
  );
  return result;
}

function runInstaller(env, expect = 0) {
  if (windows) {
    return execute(
      path.join(process.env.SystemRoot, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", installer],
      { env, expect },
    );
  }
  return execute("sh", [installer], { env, expect });
}

async function countTemplateCards(packageRoot) {
  let count = 0;
  for (const catalog of ["default-template-library", "presentation-template-library"]) {
    const skillsRoot = path.join(packageRoot, "skills", catalog, "skills");
    const children = await fs.readdir(skillsRoot, { withFileTypes: true });
    for (const child of children) {
      if (!child.isDirectory() || !child.name.startsWith("artifact-template-")) continue;
      await fs.access(path.join(skillsRoot, child.name, "artifact-template.json"));
      count += 1;
    }
  }
  return count;
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-standalone-"));
try {
  if (!windows) {
    const compressionFixture = path.join(temporary, "compression-fixture");
    await fs.mkdir(path.join(compressionFixture, "sub"), { recursive: true });
    await fs.writeFile(path.join(compressionFixture, "alpha.txt"), "alpha\n");
    await fs.writeFile(
      path.join(compressionFixture, "sub", "run"),
      "#!/bin/sh\nexit 0\n",
      { mode: 0o755 },
    );
    await fs.chmod(path.join(compressionFixture, "sub", "run"), 0o755);
    const compressionVector = await createDeterministicTarGz(
      compressionFixture,
      "fixture",
    );
    assert.equal(compressionVector.length, 174);
    assert.equal(
      sha256(compressionVector),
      "e444435e9be092e8a177e1f8c448c101c4cf8b453e8a89a956e8752b1553ea13",
      "release compression must stay independent of the host Node/zlib version",
    );
  }

  const fakeRuntime = path.join(temporary, "fake-runtime");
  const runtimeRootName = `node-v${process.versions.node}-${windows ? "win-x64" : target}`;
  const runtimeRoot = path.join(fakeRuntime, runtimeRootName);
  if (windows) {
    await fs.mkdir(runtimeRoot, { recursive: true });
    await fs.copyFile(process.execPath, path.join(runtimeRoot, "node.exe"));
  } else {
    await fs.mkdir(path.join(runtimeRoot, "bin"), { recursive: true });
    await fs.writeFile(
      path.join(runtimeRoot, "bin", "node"),
      `#!/bin/sh\nexec ${JSON.stringify(process.execPath)} "$@"\n`,
      "utf8",
    );
    await fs.chmod(path.join(runtimeRoot, "bin", "node"), 0o755);
  }
  const nodeLicenseCandidates = [
    path.join(path.dirname(process.execPath), "LICENSE"),
    path.resolve(path.dirname(process.execPath), "..", "LICENSE"),
  ];
  let nodeLicense = null;
  for (const candidate of nodeLicenseCandidates) {
    try {
      await fs.access(candidate);
      nodeLicense = candidate;
      break;
    } catch {
      // Node's Windows archive keeps LICENSE alongside node.exe; Unix layouts
      // commonly keep it one directory above the executable.
    }
  }
  assert.ok(nodeLicense, "the local Node runtime must include a LICENSE file");
  await fs.copyFile(nodeLicense, path.join(runtimeRoot, "LICENSE"));
  const runtimeArchive = path.join(temporary, `${runtimeRootName}${windows ? ".zip" : ".tar.gz"}`);
  const runtimeArchiveBytes = windows
    ? await createDeterministicZip(runtimeRoot, runtimeRootName)
    : await createDeterministicTarGz(runtimeRoot, runtimeRootName);
  await fs.writeFile(runtimeArchive, runtimeArchiveBytes);
  const runtimeEntry = {
    archive: path.basename(runtimeArchive),
    root: runtimeRootName,
    url: `https://nodejs.org/dist/v${process.versions.node}/${path.basename(runtimeArchive)}`,
    sha256: sha256(runtimeArchiveBytes),
    size: runtimeArchiveBytes.length,
  };

  const outputA = path.join(temporary, "release-a");
  const outputB = path.join(temporary, "release-b");
  const first = await buildStandalone({
    target,
    outputDirectory: outputA,
    runtimeArchive,
    runtimeEntry,
    nodeVersion: process.versions.node,
  });
  const second = await buildStandalone({
    target,
    outputDirectory: outputB,
    runtimeArchive,
    runtimeEntry,
    nodeVersion: process.versions.node,
  });
  assert.equal(first.metadata.sha256, second.metadata.sha256);
  assert.equal(first.metadata.size, second.metadata.size);
  assert.equal(
    sha256(await fs.readFile(first.archive)),
    first.metadata.sha256,
    "release metadata must describe the emitted archive",
  );
  assert.equal(first.metadata.officeKitVersion, packageMetadata.version);
  assert.equal(first.metadata.target, target);
  assert.ok(first.metadata.unpackedBytes > 40_000_000);
  assert.ok(first.metadata.fileCount > 600);
  await assert.rejects(
    buildStandalone({
      target,
      outputDirectory: outputA,
      runtimeArchive,
      runtimeEntry,
      nodeVersion: process.versions.node,
    }),
    /refusing to overwrite/,
  );
  await assert.rejects(
    buildStandalone({
      target,
      outputDirectory: path.join(temporary, "bad-runtime"),
      runtimeArchive,
      runtimeEntry: { ...runtimeEntry, sha256: "0".repeat(64) },
      nodeVersion: process.versions.node,
    }),
    /SHA-256/,
  );
  assert.equal(
    sha256(await fs.readFile(first.sbom)),
    first.metadata.sbom.sha256,
  );
  assert.equal(
    sha256(await fs.readFile(first.notices)),
    first.metadata.notices.sha256,
  );

  const home = path.join(temporary, "home");
  const installRoot = path.join(home, ".office-kit");
  const binRoot = windows
    ? path.join(installRoot, "bin")
    : path.join(home, ".local", "bin");
  const trapBin = path.join(temporary, "trap-bin");
  const trapLog = path.join(temporary, "unexpected-system-runtime.log");
  await fs.mkdir(trapBin);
  for (const command of ["node", "npm", "npx"]) {
    const trap = path.join(trapBin, windows ? `${command}.cmd` : command);
    await fs.writeFile(
      trap,
      windows
        ? `@echo off\necho ${command}>>${JSON.stringify(trapLog)}\nexit /b 99\n`
        : `#!/bin/sh\nprintf '%s\\n' ${command} >> ${JSON.stringify(trapLog)}\nexit 99\n`,
      { mode: 0o755 },
    );
    if (!windows) await fs.chmod(trap, 0o755);
  }
  const runtimePath = windows
    ? `${trapBin};${path.join(process.env.SystemRoot, "System32")}`
    : `${trapBin}:/usr/bin:/bin`;
  const installEnvironment = {
    HOME: home,
    PATH: runtimePath,
    OFFICE_KIT_HOME: installRoot,
    OFFICE_KIT_BIN_DIR: binRoot,
    OFFICE_KIT_INSTALL_TEST: "1",
    OFFICE_KIT_TEST_VERSION: packageMetadata.version,
    OFFICE_KIT_TEST_TARGET: target,
    OFFICE_KIT_TEST_ARCHIVE: first.archive,
    OFFICE_KIT_TEST_SHA256: first.metadata.sha256,
    OFFICE_KIT_TEST_SIZE: String(first.metadata.size),
    SHELL: "/bin/zsh",
    OFFICE_KIT_TEST_CONFIGURE_PATH: windows ? "0" : "1",
  };
  runInstaller(installEnvironment);

  const officekit = path.join(binRoot, windows ? "officekit.cmd" : "officekit");
  if (!windows) {
    const profile = await fs.readFile(path.join(home, ".zshrc"), "utf8");
    assert.ok(profile.includes(`export PATH="${binRoot}:$PATH"`));
  }
  const activeRoot = windows
    ? path.join(
      installRoot,
      "versions",
      (await fs.readFile(path.join(installRoot, "current.version"), "utf8")).trim(),
    )
    : path.join(installRoot, "current");
  const installedPackage = path.join(
    activeRoot,
    "app",
    "node_modules",
    "office-kit",
  );
  const runOfficeKit = (args, cwd, env = {}, expect = 0) =>
    execute(officekit, args, {
      cwd,
      env: { ...installEnvironment, ...env },
      expect,
    });

  const versionProbe = runOfficeKit(["--version"], temporary, {
    NODE_DEBUG: "esm",
  });
  assert.equal(versionProbe.stdout.trim(), packageMetadata.version);
  assert.doesNotMatch(
    versionProbe.stderr,
    /node_modules\/mupdf|src\/pdf\/mupdf|runtime\/office-kit\/main|src\/codecs\//iu,
    "version probe must not initialize Office or PDF runtimes",
  );

  const manifest = JSON.parse(
    await fs.readFile(path.join(activeRoot, "standalone-manifest.json"), "utf8"),
  );
  assert.equal(manifest.schema, "office-kit.standalone.v1");
  assert.equal(manifest.officeKitVersion, packageMetadata.version);
  assert.equal(manifest.target, target);
  assert.ok(
    manifest.files.some((entry) => entry.path === (windows
      ? "runtime/node/node.exe"
      : "runtime/node/bin/node")),
  );
  assert.ok(
    manifest.files.some((entry) => entry.path === "lib/verify-install.mjs"),
  );
  assert.ok(
    manifest.files.some(
      (entry) => entry.path === `app/node_modules/office-kit-codec-${target}/manifest.json`,
    ),
  );
  const sbom = JSON.parse(
    await fs.readFile(path.join(activeRoot, "sbom.cdx.json"), "utf8"),
  );
  assert.equal(sbom.bomFormat, "CycloneDX");
  assert.ok(sbom.components.some((component) => component.name === "Node.js"));
  assert.ok(sbom.components.some((component) => component.name === "office-kit"));
  assert.ok(sbom.components.some((component) => component.name === "mupdf"));
  assert.ok(sbom.components.some((component) => component.name === "@firecrawl/anydoc"));
  assert.ok(sbom.components.some((component) => component.name === anydocNativePackage));
  assert.ok(sbom.components.some((component) => component.name === "setimmediate"));
  await fs.access(path.join(activeRoot, "app", "node_modules", "setimmediate", "package.json"));
  await fs.access(path.join(activeRoot, "licenses", "OFFICEKIT-LICENSE.txt"));
  await fs.access(path.join(activeRoot, "licenses", "NODE-LICENSE.txt"));
  assert.equal(await countTemplateCards(installedPackage), 52);

  const project = path.join(temporary, "empty-project");
  await fs.mkdir(project);
  const initializationProbe = runOfficeKit(
    ["init", ".", "--tools", "agents", "--json"],
    project,
    { NODE_DEBUG: "esm" },
  );
  const initialized = JSON.parse(initializationProbe.stdout);
  assert.equal(initialized.created, 11);
  assert.doesNotMatch(
    initializationProbe.stderr,
    /node_modules\/mupdf|src\/pdf\/mupdf|runtime\/office-kit\/main|src\/codecs\//iu,
    "init must not initialize Office or PDF runtimes",
  );
  assert.equal(
    await fs
      .access(path.join(project, "node_modules"))
      .then(() => true, () => false),
    false,
  );
  let templateCount = 0;
  for (const [kind, expected] of [
    ["document", 7],
    ["spreadsheet", 6],
    ["presentation", 20],
  ]) {
    const searchProbe = runOfficeKit(
      ["template", "search", "--kind", kind, "--max", "20", "--json"],
      project,
      { NODE_DEBUG: "esm" },
    );
    const search = JSON.parse(searchProbe.stdout);
    assert.equal(search.candidates.length, expected);
    assert.equal(search.selectionMade, false);
    assert.deepEqual(search.invalid, []);
    assert.ok(search.searchedRoots.some((root) => root.source === "package-default"));
    assert.doesNotMatch(
      searchProbe.stderr,
      /node_modules\/mupdf|src\/pdf\/mupdf|runtime\/office-kit\/main|src\/codecs\//iu,
      "template search must not initialize Office or PDF runtimes",
    );
    templateCount += search.candidates.length;
  }
  assert.equal(templateCount, 33);
  assert.equal(
    await fs
      .access(path.join(project, ".open-office-artifact-tool", "providers"))
      .then(() => true, () => false),
    false,
  );

  const task = path.join(
    repositoryRoot,
    "test",
    "fixtures",
    "standalone-four-formats.mjs",
  );
  const taskResult = JSON.parse(
    runOfficeKit(["run", task, "--", "alpha"], project).stdout,
  );
  assert.deepEqual(taskResult.argv, ["alpha"]);
  assert.equal(await fs.realpath(taskResult.cwd), await fs.realpath(project));
  assert.equal(taskResult.publicSubpaths, Object.keys(packageMetadata.exports).length);
  assert.equal(taskResult.anydoc, "ready");
  const standaloneProgram = JSON.parse(await fs.readFile(
    path.join(installedPackage, "examples", "ppj", "minimum.ppj"),
    "utf8",
  ));
  standaloneProgram.meta.id = "standalone";
  standaloneProgram.meta.title = "Standalone";
  standaloneProgram.pages[0].id = "page-1";
  standaloneProgram.pages[0].elements[0].id = "title";
  standaloneProgram.pages[0].elements[0].text = "standalone PPTX";
  await fs.writeFile(
    path.join(project, "standalone.ppj"),
    `${JSON.stringify(standaloneProgram, null, 2)}\n`,
  );
  const ppjBuild = JSON.parse(runOfficeKit([
    "ppj", "build", "standalone.ppj", "-o", "standalone.pptx", "--json",
  ], project).stdout);
  assert.equal(ppjBuild.ok, true);
  for (const extension of ["docx", "xlsx", "pptx", "pdf"]) {
    assert.ok((await fs.stat(path.join(project, `standalone.${extension}`))).size > 100);
  }

  const recovered = JSON.parse(runOfficeKit([
    "ppj", "import", "standalone.pptx", "-o", "standalone-recovered.ppj", "--json",
  ], project).stdout);
  assert.equal(recovered.ok, true);
  const recoveredProgram = JSON.parse(await fs.readFile(path.join(project, "standalone-recovered.ppj"), "utf8"));
  assert.equal(recoveredProgram.pages[0]?.elements[0]?.id, "title");
  assert.equal(recoveredProgram.pages[0]?.elements[0]?.text, "standalone PPTX");

  const dependencyProject = path.join(temporary, "task-local-dependency");
  const dependencyRoot = path.join(
    dependencyProject,
    "node_modules",
    "local-probe",
  );
  await fs.mkdir(dependencyRoot, { recursive: true });
  await fs.writeFile(
    path.join(dependencyRoot, "package.json"),
    `${JSON.stringify({
      name: "local-probe",
      version: "1.1.0",
      type: "module",
      exports: "./index.mjs",
    })}\n`,
  );
  await fs.writeFile(
    path.join(dependencyRoot, "index.mjs"),
    "export default 41;\n",
  );
  await fs.writeFile(
    path.join(dependencyProject, "task.mjs"),
    'import value from "local-probe"; console.log(value + 1);\n',
  );
  assert.equal(
    runOfficeKit(["run", "task.mjs"], dependencyProject).stdout.trim(),
    "42",
  );

  assert.equal(
    await fs.readFile(trapLog, "utf8").catch(() => ""),
    "",
    "installer and installed command must not invoke system node, npm, or npx",
  );

  const currentState = async () => windows
    ? fs.readFile(path.join(installRoot, "current.version"), "utf8")
    : fs.readlink(activeRoot);
  const activeBeforeFailure = await currentState();
  const rejected = runInstaller({
    ...installEnvironment,
    OFFICE_KIT_TEST_SHA256: "0".repeat(64),
  }, 1);
  assert.match(rejected.stderr, /archive SHA-256/);
  assert.equal(await currentState(), activeBeforeFailure);
  assert.equal(runOfficeKit(["--version"], temporary).stdout.trim(), packageMetadata.version);

  const installedReadme = path.join(installedPackage, "README.md");
  const originalReadme = await fs.readFile(installedReadme);
  await fs.appendFile(installedReadme, "\nlocal corruption fixture\n");
  const corrupted = runInstaller(installEnvironment, 1);
  assert.match(corrupted.stderr, /file integrity verification failed/);
  assert.equal(await currentState(), activeBeforeFailure);
  await fs.writeFile(installedReadme, originalReadme);

  runInstaller(installEnvironment);
  assert.equal(await currentState(), activeBeforeFailure);
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("self-contained OfficeKit distribution smoke ok");
