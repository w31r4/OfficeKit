import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { EventEmitter } from "node:events";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { Readable } from "node:stream";

import JSZip from "jszip";

import { runImageCommand } from "../src/images/cli.mjs";
import { downloadRemoteImage } from "../src/images/download.mjs";
import { searchImageCandidates } from "../src/images/providers.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const cli = path.join(repoRoot, "bin", "officekit.mjs");
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "officekit-cli-"));
const lazyExcelHome = path.join(temporary, "lazy-excel-state");
const lazyPowerPointHome = path.join(temporary, "lazy-powerpoint-state");

try {
  const help = run(["--help"]);
  assert.match(help.stdout, /officekit init \[path\]/);
  assert.match(help.stdout, /officekit update \[path\]/);
  assert.match(help.stdout, /officekit run <task\.mjs>/);
  assert.match(help.stdout, /officekit tasks \[<task-id>\]/);
  assert.match(help.stdout, /officekit repl --new <goal>/);
  assert.match(help.stdout, /officekit repl <task-id>/);
  assert.match(run(["repl", "--help"]).stdout, /--file <cell[.]mjs>/);
  assert.match(help.stdout, /officekit template search/);
  assert.match(help.stdout, /officekit image <search\|add\|list\|audit>/);
  assert.match(help.stdout, /officekit excel <command>/);
  assert.match(help.stdout, /officekit live <command> --app <excel\|powerpoint>/);
  assert.match(help.stdout, /Choose Agent targets and install the OfficeKit Skills/);
  assert.equal(run(["--version"]).stdout.trim(), "1.1.0");
  const excelHelp = run(["excel", "--help"]);
  assert.match(excelHelp.stdout, /officekit excel install/);
  assert.match(excelHelp.stdout, /officekit excel execute <request\.json>/);
  const liveHelp = run(["live", "--help"]);
  assert.match(liveHelp.stdout, /officekit live install --app powerpoint/);
  assert.match(liveHelp.stdout, /officekit live execute <request\.json>/);
  assert.match(run(["image", "--help"]).stdout, /selectionMade is always false/);

  const lazyProject = path.join(temporary, "lazy-excel-project");
  const lazyEnvironment = { OFFICEKIT_EXCEL_HOME: lazyExcelHome, OFFICEKIT_POWERPOINT_HOME: lazyPowerPointHome };
  parseJson(run(["init", lazyProject, "--tools", "agents", "--json"], { environment: lazyEnvironment }).stdout);
  const ready = parseJson(run(["repl", "--new", "CLI discovery task", "--workspace", lazyProject], { environment: lazyEnvironment }).stdout);
  assert.equal(ready.type, "session.ready");
  assert.equal(ready.task.goal, "CLI discovery task");
  const tasks = parseJson(run(["tasks", "--workspace", lazyProject, "--json"], { environment: lazyEnvironment }).stdout);
  assert.equal(tasks.total, 1);
  assert.equal(tasks.tasks[0].id, ready.task.id);
  const taskDetail = parseJson(run(["tasks", ready.task.id, "--workspace", lazyProject, "--json"], { environment: lazyEnvironment }).stdout);
  assert.equal(taskDetail.task.goal, "CLI discovery task");

  let directSearchOutput = "";
  const directSearch = await runImageCommand([
    "search", "market evidence", "--task", ready.task.id, "--workspace", lazyProject,
    "--kind", "photo", "--purpose", "evidence", "--orientation", "landscape", "--max", "2", "--json",
  ], {
    output: { write(chunk) { directSearchOutput += chunk; } },
    searcher: (input) => searchImageCandidates(input, {
      providerImplementations: {
        openverse: { search: async () => [
          { url: "https://images.example.com/market.png", sourcePageUrl: "https://example.com/market", title: "Market evidence", author: "Example Author", license: "CC_BY", licenseUrl: "https://creativecommons.org/licenses/by/4.0/", width: 1600, height: 900, mime: "image/png" },
          { url: "https://images.example.com/blocked.png", sourcePageUrl: "https://example.com/blocked", title: "Blocked", author: "Example Author", license: "CC_BY_SA", width: 1600, height: 900, mime: "image/png" },
        ] },
        wikimedia: { search: async () => [] },
      },
    }),
  });
  assert.equal(directSearch.selectionMade, false);
  assert.equal(directSearch.candidates.length, 1);
  assert.equal(directSearch.candidates[0].rights, "cc-by");
  assert.equal(directSearch.rejected[0].reason, "image-rights-blocked");
  assert.equal(parseJson(directSearchOutput).selectionMade, false);

  const iconSearch = parseJson(run([
    "image", "search", "market chart", "--task", ready.task.id, "--workspace", lazyProject,
    "--kind", "icon", "--purpose", "context", "--orientation", "square", "--max", "1", "--json",
  ]).stdout);
  assert.equal(iconSearch.selectionMade, false);
  assert.equal(iconSearch.candidates.length, 1);
  const iconAsset = parseJson(run([
    "image", "add", "--task", ready.task.id, "--workspace", lazyProject,
    "--candidate", iconSearch.candidates[0].candidateRef, "--json",
  ]).stdout).asset;
  assert.equal(iconAsset.rights, "lucide-isc");

  const pngBytes = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64");
  const localImage = path.join(lazyProject, "evidence.png");
  fs.writeFileSync(localImage, pngBytes);
  const localAsset = parseJson(run([
    "image", "add", "--task", ready.task.id, "--workspace", lazyProject,
    "--file", localImage, "--rights", "user-provided", "--json",
  ]).stdout).asset;
  assert.equal(localAsset.mimeType, "image/png");
  assert.equal(parseJson(run(["image", "list", "--task", ready.task.id, "--workspace", lazyProject, "--json"]).stdout).assets.length, 2);

  const pptxZip = new JSZip();
  pptxZip.file("ppt/media/image1.png", pngBytes);
  const auditPptx = path.join(lazyProject, "image-audit.pptx");
  fs.writeFileSync(auditPptx, await pptxZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
  const sourcesOutput = `${auditPptx}.sources.json`;
  const imageAudit = parseJson(run([
    "image", "audit", auditPptx, "--task", ready.task.id, "--workspace", lazyProject,
    "--sources-output", sourcesOutput, "--json",
  ]).stdout);
  assert.equal(imageAudit.audit.ok, true);
  assert.equal(imageAudit.audit.used[0].sha256, localAsset.sha256);
  assert.equal(imageAudit.audit.unused[0].sha256, iconAsset.sha256);
  assert.equal(fs.existsSync(sourcesOutput), true);
  assert.match(run([
    "image", "audit", auditPptx, "--task", ready.task.id, "--workspace", lazyProject,
    "--sources-output", sourcesOutput, "--json",
  ], { expectFailure: true }).stderr, /image-output-exists/);

  const publicResolver = async () => [{ address: "93.184.216.34", family: 4 }];
  const redirected = await downloadRemoteImage("https://images.example.com/start", {
    resolver: publicResolver,
    requestFactory: queuedRequestFactory([
      response(302, { location: "https://cdn.example.com/final.png" }),
      response(200, { "content-type": "image/png", "content-length": String(pngBytes.length) }, [pngBytes]),
    ]),
  });
  assert.equal(redirected.redirects.length, 1);
  await assert.rejects(
    downloadRemoteImage("https://images.example.com/private.png", { resolver: async () => [{ address: "127.0.0.1", family: 4 }] }),
    (error) => error.code === "unsafe-image-destination",
  );
  await assert.rejects(
    downloadRemoteImage("https://images.example.com/large.png", { resolver: publicResolver, requestFactory: queuedRequestFactory([response(200, { "content-type": "image/png", "content-length": String(21 * 1024 * 1024) })]) }),
    (error) => error.code === "image-download-too-large",
  );
  await assert.rejects(
    downloadRemoteImage("https://images.example.com/wrong.png", { resolver: publicResolver, requestFactory: queuedRequestFactory([response(200, { "content-type": "image/png" }, [Buffer.from("not a png")])]) }),
    /PNG|image/i,
  );
  parseJson(run([
    "template",
    "search",
    "--kind",
    "document",
    "--purpose",
    "board briefing",
    "--json",
  ], { cwd: lazyProject, environment: lazyEnvironment }).stdout);
  assert.equal(
    fs.existsSync(lazyExcelHome),
    false,
    "root CLI initialization and template search must not initialize the Excel bridge or state",
  );
  assert.equal(
    fs.existsSync(lazyPowerPointHome),
    false,
    "root CLI initialization and template search must not initialize the PowerPoint bridge or state",
  );
  assert.match(
    run(["tasks", "--delete", ready.task.id, "--workspace", lazyProject, "--json"], { expectFailure: true }).stderr,
    /requires --yes/,
  );
  const deletedTask = parseJson(run(["tasks", "--delete", ready.task.id, "--yes", "--workspace", lazyProject, "--json"]).stdout);
  assert.equal(deletedTask.deleted, true);
  assert.equal(parseJson(run(["tasks", "--workspace", lazyProject, "--json"]).stdout).total, 0);

  const project = path.join(temporary, "detected-project");
  fs.mkdirSync(path.join(project, ".claude"), { recursive: true });
  fs.mkdirSync(path.join(project, ".cursor"), { recursive: true });
  const initialized = parseJson(run(["init", project, "--yes", "--json"]).stdout);
  assert.equal(initialized.ok, true);
  assert.deepEqual(initialized.tools.map((tool) => tool.id), ["claude", "cursor"]);
  assert.equal(initialized.created, 18);
  assert.equal(initialized.updated, 0);
  assert.equal(initialized.unchanged, 0);
  for (const toolRoot of [".claude", ".cursor"]) {
    for (const skill of [
      "office-kit",
      "documents",
      "spreadsheets",
      "excel-live-control",
      "presentations",
      "presentation-editorial-trim",
      "powerpoint-live-control",
      "pdf",
      "template-creator",
    ]) {
      assert.ok(fs.existsSync(path.join(project, toolRoot, "skills", skill, "SKILL.md")));
    }
    assert.ok(
      fs.existsSync(
        path.join(project, toolRoot, "skills", "template-creator", "assets", "icon.svg"),
      ),
    );
  }

  const manifestPath = path.join(project, ".office-kit", "skills.json");
  const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
  assert.equal(manifest.schemaVersion, 1);
  assert.deepEqual(manifest.tools, ["claude", "cursor"]);
  assert.equal(manifest.installations.length, 18);
  assert.equal(manifest.package.name, "office-kit");
  assert.equal(manifest.package.version, "1.1.0");

  const idempotent = parseJson(run(["init", project, "--yes", "--json"]).stdout);
  assert.equal(idempotent.created, 0);
  assert.equal(idempotent.updated, 0);
  assert.equal(idempotent.unchanged, 18);

  const managedSkill = path.join(project, ".claude", "skills", "office-kit", "SKILL.md");
  const sourceSkill = path.join(
    repoRoot,
    "skills",
    "office-kit",
    "skills",
    "office-kit",
    "SKILL.md",
  );
  fs.appendFileSync(managedSkill, "\nlocal edit\n");
  const guarded = run(["update", project, "--json"], { expectFailure: true });
  assert.match(guarded.stderr, /changed after OfficeKit installed it/);
  assert.match(fs.readFileSync(managedSkill, "utf8"), /local edit/);

  const restored = parseJson(
    run(["update", project, "--force", "--json"]).stdout,
  );
  assert.equal(restored.updated, 1);
  assert.equal(restored.unchanged, 17);
  assert.equal(
    sha256(fs.readFileSync(managedSkill)),
    sha256(fs.readFileSync(sourceSkill)),
  );

  const collisionProject = path.join(temporary, "collision-project");
  const collisionSkill = path.join(
    collisionProject,
    ".agents",
    "skills",
    "office-kit",
    "SKILL.md",
  );
  fs.mkdirSync(path.dirname(collisionSkill), { recursive: true });
  fs.writeFileSync(collisionSkill, "user-owned\n");
  const collision = run(
    ["init", collisionProject, "--tools", "agents", "--force", "--json"],
    { expectFailure: true },
  );
  assert.match(collision.stderr, /not managed by OfficeKit/);
  assert.equal(fs.readFileSync(collisionSkill, "utf8"), "user-owned\n");
  assert.equal(
    fs.existsSync(path.join(collisionProject, ".agents", "skills", "documents")),
    false,
    "preflight must reject a collision before writing another Skill",
  );

  const identicalCollisionProject = path.join(temporary, "identical-collision-project");
  const identicalCollisionSkill = path.join(
    identicalCollisionProject,
    ".agents",
    "skills",
    "office-kit",
  );
  fs.cpSync(
    path.join(repoRoot, "skills", "office-kit", "skills", "office-kit"),
    identicalCollisionSkill,
    { recursive: true },
  );
  const identicalCollision = run(
    ["init", identicalCollisionProject, "--tools", "agents", "--force", "--json"],
    { expectFailure: true },
  );
  assert.match(identicalCollision.stderr, /not managed by OfficeKit/);
  assert.equal(
    fs.existsSync(path.join(identicalCollisionProject, ".office-kit")),
    false,
    "matching bytes must not grant OfficeKit ownership without a manifest",
  );

  const explicitProject = path.join(temporary, "explicit-project");
  const explicit = parseJson(
    run(["init", explicitProject, "--tools", "agents", "--json"]).stdout,
  );
  assert.deepEqual(explicit.tools.map((tool) => tool.id), ["agents"]);
  assert.ok(
    fs.existsSync(
      path.join(explicitProject, ".agents", "skills", "office-kit", "SKILL.md"),
    ),
  );
  assert.equal(
    fs.existsSync(path.join(explicitProject, "skills", "default-template-library")),
    false,
  );
  const catalog = parseJson(
    run(
      [
        "template",
        "search",
        "--kind",
        "presentation",
        "--purpose",
        "quarterly business review",
        "--max",
        "20",
        "--json",
      ],
      { cwd: explicitProject },
    ).stdout,
  );
  assert.equal(catalog.selectionMade, false);
  assert.equal(catalog.candidates[0].id, "artifact-template-business-review");
  assert.ok(catalog.searchedRoots.some((root) => root.source === "project"));
  assert.ok(catalog.searchedRoots.some((root) => root.source === "package-default"));
  const catalogTable = run(
    [
      "template",
      "search",
      "--kind",
      "presentation",
      "--purpose",
      "quarterly business review",
    ],
    { cwd: explicitProject },
  ).stdout;
  assert.match(catalogTable, /Rank\s+Template\s+Score\s+Coverage\s+Review/);
  assert.match(catalogTable, /artifact-template-business-review/);
  assert.match(catalogTable, /Selection remains with the Agent/);

  const noCatalogMatch = parseJson(
    run(
      [
        "template",
        "search",
        "--kind",
        "presentation",
        "--purpose",
        "quantum entanglement laboratory protocol",
        "--json",
      ],
      { cwd: explicitProject },
    ).stdout,
  );
  assert.equal(noCatalogMatch.retrievalStatus, "none");
  assert.deepEqual(noCatalogMatch.candidates, []);
  assert.equal(noCatalogMatch.selectionMade, false);

  const taskProject = path.join(temporary, "run-project");
  const dependencyRoot = path.join(taskProject, "node_modules", "local-probe");
  fs.mkdirSync(dependencyRoot, { recursive: true });
  fs.writeFileSync(
    path.join(dependencyRoot, "package.json"),
    `${JSON.stringify({
      name: "local-probe",
      version: "1.0.0",
      type: "module",
      exports: "./index.mjs",
    })}\n`,
  );
  fs.writeFileSync(
    path.join(dependencyRoot, "index.mjs"),
    "export default 'resolved-from-task-project';\n",
  );
  const taskPath = path.join(taskProject, "task.mjs");
  fs.writeFileSync(
    taskPath,
    [
      'import { createRequire } from "node:module";',
      'import { FileBlob } from "office-kit";',
      'import * as wire from "office-kit/codec/wire";',
      'import localProbe from "local-probe";',
      "const require = createRequire(import.meta.url);",
      "console.log(JSON.stringify({",
      "  argv: process.argv.slice(2),",
      "  cwd: process.cwd(),",
      "  fileBlob: typeof FileBlob === 'function',",
      "  wire: wire.CodecRequestSchema != null,",
      '  resolvedOfficeKit: require.resolve("office-kit"),',
      "  localProbe,",
      "}));",
      "",
    ].join("\n"),
  );
  const taskResult = parseJson(
    run(["run", "task.mjs", "--", "alpha", "two words"], {
      cwd: taskProject,
    }).stdout,
  );
  assert.deepEqual(taskResult.argv, ["alpha", "two words"]);
  assert.equal(taskResult.cwd, fs.realpathSync(taskProject));
  assert.equal(taskResult.fileBlob, true);
  assert.equal(taskResult.wire, true);
  assert.equal(
    path.resolve(taskResult.resolvedOfficeKit),
    path.join(repoRoot, "src", "index.mjs"),
  );
  assert.equal(taskResult.localProbe, "resolved-from-task-project");
  assert.equal(
    fs.existsSync(path.join(taskProject, "node_modules", "office-kit")),
    false,
    "officekit run must not require a project-local OfficeKit package",
  );

  const failedTaskPath = path.join(taskProject, "failed-task.mjs");
  fs.writeFileSync(failedTaskPath, 'throw new Error("task stack sentinel");\n');
  const failedTask = run(["run", failedTaskPath], {
    cwd: taskProject,
    expectFailure: true,
  });
  assert.match(failedTask.stderr, /task stack sentinel/);
  assert.match(failedTask.stderr, /failed-task\.mjs:1/);
  assert.doesNotMatch(failedTask.stderr, /^OfficeKit:/);

  const privateSubpathPath = path.join(taskProject, "private-subpath.mjs");
  fs.writeFileSync(
    privateSubpathPath,
    'await import("office-kit/src/index.mjs");\n',
  );
  const privateSubpath = run(["run", privateSubpathPath], {
    cwd: taskProject,
    expectFailure: true,
  });
  assert.match(privateSubpath.stderr, /unpublished package subpath/);
  assert.match(
    run(["run", "https://example.com/task.mjs"], {
      cwd: taskProject,
      expectFailure: true,
    }).stderr,
    /not stdin or a URL/,
  );

  const invalid = run(
    ["init", path.join(temporary, "invalid"), "--tools", "unknown", "--json"],
    { expectFailure: true },
  );
  assert.match(invalid.stderr, /Unknown Agent tool/);

  const noDetectedTarget = run(
    ["init", path.join(temporary, "no-agent-target"), "--yes", "--json"],
    { expectFailure: true },
  );
  assert.match(noDetectedTarget.stderr, /interactive terminal to choose a target/);
  assert.match(noDetectedTarget.stderr, /--tools codex/);

  const uninitializedUpdate = run(
    ["update", path.join(temporary, "not-initialized"), "--tools", "agents", "--json"],
    { expectFailure: true },
  );
  assert.match(uninitializedUpdate.stderr, /not initialized/);

  if (process.platform !== "win32") {
    const symlinkProject = path.join(temporary, "symlink-project");
    const outside = path.join(temporary, "outside");
    fs.mkdirSync(path.join(symlinkProject, ".cursor"), { recursive: true });
    fs.mkdirSync(outside, { recursive: true });
    fs.symlinkSync(outside, path.join(symlinkProject, ".cursor", "skills"), "dir");
    const symlinked = run(
      ["init", symlinkProject, "--tools", "cursor", "--json"],
      { expectFailure: true },
    );
    assert.match(symlinked.stderr, /must be a regular directory/);
    assert.deepEqual(fs.readdirSync(outside), []);

    const manifestSymlinkProject = path.join(temporary, "manifest-symlink-project");
    const manifestOutside = path.join(temporary, "manifest-outside");
    fs.mkdirSync(manifestSymlinkProject, { recursive: true });
    fs.mkdirSync(manifestOutside, { recursive: true });
    fs.symlinkSync(
      manifestOutside,
      path.join(manifestSymlinkProject, ".office-kit"),
      "dir",
    );
    const manifestSymlink = run(
      ["init", manifestSymlinkProject, "--tools", "agents", "--json"],
      { expectFailure: true },
    );
    assert.match(manifestSymlink.stderr, /\.office-kit must be a regular directory/);
    assert.equal(
      fs.existsSync(path.join(manifestSymlinkProject, ".agents")),
      false,
      "manifest-root preflight must run before Skill placement",
    );
  }
} finally {
  fs.rmSync(temporary, { recursive: true, force: true });
}

console.log("OfficeKit CLI initialization smoke ok");

function run(args, {
  expectFailure = false,
  cwd = repoRoot,
  environment = {},
} = {}) {
  const result = spawnSync(process.execPath, [cli, ...args], {
    cwd,
    encoding: "utf8",
    env: { ...process.env, ...environment },
    shell: false,
  });
  if (expectFailure) {
    assert.notEqual(result.status, 0, `officekit ${args.join(" ")} unexpectedly passed`);
  } else {
    assert.equal(
      result.status,
      0,
      `officekit ${args.join(" ")} failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`,
    );
  }
  return result;
}

function parseJson(source) {
  return JSON.parse(source.trim());
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function response(statusCode, headers, chunks = []) {
  return Object.assign(Readable.from(chunks), { statusCode, headers });
}

function queuedRequestFactory(responses) {
  return (_url, _options, callback) => {
    const request = new EventEmitter();
    request.setTimeout = () => request;
    request.destroy = (error) => queueMicrotask(() => request.emit("error", error));
    request.end = () => queueMicrotask(() => callback(responses.shift()));
    return request;
  };
}
