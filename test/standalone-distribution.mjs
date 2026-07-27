import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  buildStandalone,
  createDeterministicTarGz,
} from "../scripts/build-standalone.mjs";

const repositoryRoot = path.resolve(import.meta.dirname, "..");
const installer = path.join(repositoryRoot, "standalone", "install.sh");
const packageMetadata = JSON.parse(
  await fs.readFile(path.join(repositoryRoot, "package.json"), "utf8"),
);

const target =
  process.platform === "darwin" && process.arch === "arm64"
    ? "darwin-arm64"
    : process.platform === "linux" && process.arch === "x64"
      ? "linux-x64"
      : null;

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
  });
  assert.equal(
    result.status,
    expect,
    `${command} ${args.join(" ")} exited ${result.status}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`,
  );
  return result;
}

async function countTemplateCards(packageRoot) {
  const skillsRoot = path.join(
    packageRoot,
    "skills",
    "default-template-library",
    "skills",
  );
  const children = await fs.readdir(skillsRoot, { withFileTypes: true });
  let count = 0;
  for (const child of children) {
    if (!child.isDirectory() || !child.name.startsWith("artifact-template-")) continue;
    await fs.access(path.join(skillsRoot, child.name, "artifact-template.json"));
    count += 1;
  }
  return count;
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-standalone-"));
try {
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

  const fakeRuntime = path.join(temporary, "fake-runtime");
  const runtimeRootName = `node-v${process.versions.node}-${target}`;
  const runtimeRoot = path.join(fakeRuntime, runtimeRootName);
  await fs.mkdir(path.join(runtimeRoot, "bin"), { recursive: true });
  await fs.writeFile(
    path.join(runtimeRoot, "bin", "node"),
    `#!/bin/sh\nexec ${JSON.stringify(process.execPath)} "$@"\n`,
    "utf8",
  );
  await fs.chmod(path.join(runtimeRoot, "bin", "node"), 0o755);
  const nodeLicense = path.resolve(path.dirname(process.execPath), "..", "LICENSE");
  await fs.copyFile(nodeLicense, path.join(runtimeRoot, "LICENSE"));
  const runtimeArchive = path.join(temporary, `${runtimeRootName}.tar.gz`);
  const runtimeArchiveBytes = await createDeterministicTarGz(
    runtimeRoot,
    runtimeRootName,
  );
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
  const binRoot = path.join(home, ".local", "bin");
  const trapBin = path.join(temporary, "trap-bin");
  const trapLog = path.join(temporary, "unexpected-system-runtime.log");
  await fs.mkdir(trapBin);
  for (const command of ["node", "npm", "npx"]) {
    const trap = path.join(trapBin, command);
    await fs.writeFile(
      trap,
      `#!/bin/sh\nprintf '%s\\n' ${command} >> ${JSON.stringify(trapLog)}\nexit 99\n`,
      { mode: 0o755 },
    );
    await fs.chmod(trap, 0o755);
  }
  const runtimePath = `${trapBin}:/usr/bin:/bin`;
  const installEnvironment = {
    HOME: home,
    PATH: runtimePath,
    OFFICE_KIT_HOME: installRoot,
    OFFICE_KIT_BIN_DIR: binRoot,
    OFFICE_KIT_INSTALL_TEST: "1",
    OFFICE_KIT_TEST_TARGET: target,
    OFFICE_KIT_TEST_ARCHIVE: first.archive,
    OFFICE_KIT_TEST_SHA256: first.metadata.sha256,
    OFFICE_KIT_TEST_SIZE: String(first.metadata.size),
  };
  execute("sh", [installer], { env: installEnvironment });

  const officekit = path.join(binRoot, "officekit");
  const activeRoot = path.join(installRoot, "current");
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
    manifest.files.some((entry) => entry.path === "runtime/node/bin/node"),
  );
  assert.ok(
    manifest.files.some((entry) => entry.path === "lib/verify-install.mjs"),
  );
  assert.ok(
    manifest.files.some(
      (entry) => entry.path === "app/node_modules/office-kit/runtime/office-kit/manifest.json",
    ),
  );
  const sbom = JSON.parse(
    await fs.readFile(path.join(activeRoot, "sbom.cdx.json"), "utf8"),
  );
  assert.equal(sbom.bomFormat, "CycloneDX");
  assert.ok(sbom.components.some((component) => component.name === "Node.js"));
  assert.ok(sbom.components.some((component) => component.name === "office-kit"));
  assert.ok(sbom.components.some((component) => component.name === "mupdf"));
  await fs.access(path.join(activeRoot, "licenses", "OFFICEKIT-LICENSE.txt"));
  await fs.access(path.join(activeRoot, "licenses", "NODE-LICENSE.txt"));
  assert.equal(await countTemplateCards(installedPackage), 20);

  const project = path.join(temporary, "empty-project");
  await fs.mkdir(project);
  const initializationProbe = runOfficeKit(
    ["init", ".", "--tools", "agents", "--json"],
    project,
    { NODE_DEBUG: "esm" },
  );
  const initialized = JSON.parse(initializationProbe.stdout);
  assert.equal(initialized.created, 7);
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
    ["presentation", 7],
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
  assert.equal(templateCount, 20);
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
  assert.equal(taskResult.cwd, await fs.realpath(project));
  assert.equal(taskResult.publicSubpaths, Object.keys(packageMetadata.exports).length);
  for (const extension of ["docx", "xlsx", "pptx", "pdf"]) {
    assert.ok((await fs.stat(path.join(project, `standalone.${extension}`))).size > 100);
  }

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
      version: "1.0.0",
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

  const activeBeforeFailure = await fs.readlink(activeRoot);
  const rejected = execute("sh", [installer], {
    env: {
      ...installEnvironment,
      OFFICE_KIT_TEST_SHA256: "0".repeat(64),
    },
    expect: 1,
  });
  assert.match(rejected.stderr, /archive SHA-256/);
  assert.equal(await fs.readlink(activeRoot), activeBeforeFailure);
  assert.equal(runOfficeKit(["--version"], temporary).stdout.trim(), packageMetadata.version);

  const installedReadme = path.join(installedPackage, "README.md");
  const originalReadme = await fs.readFile(installedReadme);
  await fs.appendFile(installedReadme, "\nlocal corruption fixture\n");
  const corrupted = execute("sh", [installer], {
    env: installEnvironment,
    expect: 1,
  });
  assert.match(corrupted.stderr, /file integrity verification failed/);
  assert.equal(await fs.readlink(activeRoot), activeBeforeFailure);
  await fs.writeFile(installedReadme, originalReadme);

  execute("sh", [installer], { env: installEnvironment });
  assert.equal(await fs.readlink(activeRoot), activeBeforeFailure);
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("self-contained OfficeKit distribution smoke ok");
