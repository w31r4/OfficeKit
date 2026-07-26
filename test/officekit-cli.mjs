import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const cli = path.join(repoRoot, "bin", "officekit.mjs");
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "officekit-cli-"));

try {
  const help = run(["--help"]);
  assert.match(help.stdout, /officekit init \[path\]/);
  assert.match(help.stdout, /officekit update \[path\]/);
  assert.equal(run(["--version"]).stdout.trim(), "0.3.0");

  const project = path.join(temporary, "detected-project");
  fs.mkdirSync(path.join(project, ".claude"), { recursive: true });
  fs.mkdirSync(path.join(project, ".cursor"), { recursive: true });
  const initialized = parseJson(run(["init", project, "--yes", "--json"]).stdout);
  assert.equal(initialized.ok, true);
  assert.deepEqual(initialized.tools.map((tool) => tool.id), ["claude", "cursor"]);
  assert.equal(initialized.created, 14);
  assert.equal(initialized.updated, 0);
  assert.equal(initialized.unchanged, 0);
  for (const toolRoot of [".claude", ".cursor"]) {
    for (const skill of [
      "office-kit",
      "documents",
      "spreadsheets",
      "excel-live-control",
      "presentations",
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
  assert.equal(manifest.installations.length, 14);
  assert.equal(manifest.package.name, "office-kit");
  assert.equal(manifest.package.version, "0.3.0");

  const idempotent = parseJson(run(["init", project, "--yes", "--json"]).stdout);
  assert.equal(idempotent.created, 0);
  assert.equal(idempotent.updated, 0);
  assert.equal(idempotent.unchanged, 14);

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
  assert.equal(restored.unchanged, 13);
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

  const invalid = run(
    ["init", path.join(temporary, "invalid"), "--tools", "unknown", "--json"],
    { expectFailure: true },
  );
  assert.match(invalid.stderr, /Unknown Agent tool/);

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

function run(args, { expectFailure = false } = {}) {
  const result = spawnSync(process.execPath, [cli, ...args], {
    cwd: repoRoot,
    encoding: "utf8",
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
