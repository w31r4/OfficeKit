import { createHash, randomUUID } from "node:crypto";
import {
  access,
  cp,
  lstat,
  mkdir,
  readFile,
  readdir,
  realpath,
  rename,
  rm,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { createInterface } from "node:readline/promises";
import { fileURLToPath } from "node:url";

const PACKAGE_ROOT = fileURLToPath(new URL("../..", import.meta.url));
const MANIFEST_PATH = ".office-kit/skills.json";

const SKILLS = Object.freeze([
  ["office-kit", "skills/office-kit/skills/office-kit"],
  ["documents", "skills/documents/skills/documents"],
  ["spreadsheets", "skills/spreadsheets/skills/spreadsheets"],
  ["excel-live-control", "skills/spreadsheets/skills/excel-live-control"],
  ["presentations", "skills/presentations/skills/presentations"],
  ["powerpoint-live-control", "skills/presentations/skills/powerpoint-live-control"],
  ["pdf", "skills/pdf/skills/pdf"],
  ["template-creator", "skills/template-creator/skills/template-creator"],
].map(([id, source]) => Object.freeze({ id, source })));

// These project-local Skill roots follow the same cross-agent layout used by
// OpenSpec. `agents` is the generic Agent Skills location and is explicit-only
// unless an existing .agents/skills directory is present.
const TOOLS = Object.freeze([
  ["agents", "Agent Skills", ".agents", [".agents/skills"]],
  ["amazon-q", "Amazon Q Developer", ".amazonq"],
  ["antigravity", "Antigravity", ".agent"],
  ["auggie", "Auggie", ".augment"],
  ["bob", "Bob Shell", ".bob"],
  ["claude", "Claude Code", ".claude"],
  ["cline", "Cline", ".cline"],
  ["codeartsagent", "CodeArts", ".codeartsdoer"],
  ["codebuddy", "CodeBuddy Code", ".codebuddy"],
  ["codex", "Codex", ".codex"],
  ["continue", "Continue", ".continue"],
  ["costrict", "CoStrict", ".cospec"],
  ["crush", "Crush", ".crush"],
  ["cursor", "Cursor", ".cursor"],
  ["factory", "Factory Droid", ".factory"],
  ["forgecode", "ForgeCode", ".forge"],
  ["gemini", "Gemini CLI", ".gemini"],
  [
    "github-copilot",
    "GitHub Copilot",
    ".github",
    [
      ".github/copilot-instructions.md",
      ".github/instructions",
      ".github/prompts",
      ".github/agents",
      ".github/skills",
      ".github/.mcp.json",
    ],
  ],
  ["hermes", "Hermes Agent", ".hermes", [".hermes", "HERMES.md", ".hermes.md"]],
  ["iflow", "iFlow", ".iflow"],
  ["junie", "Junie", ".junie"],
  ["kilocode", "Kilo Code", ".kilocode"],
  ["kimi", "Kimi Code", ".kimi-code", [".kimi-code", ".kimi"]],
  ["kiro", "Kiro", ".kiro"],
  ["lingma", "Lingma", ".lingma"],
  ["oh-my-pi", "Oh My Pi", ".omp"],
  ["opencode", "OpenCode", ".opencode"],
  ["pi", "Pi", ".pi"],
  ["qoder", "Qoder", ".qoder"],
  ["qwen", "Qwen Code", ".qwen"],
  ["roocode", "Roo Code", ".roo"],
  ["trae", "Trae", ".trae"],
  ["vibe", "Mistral Vibe", ".vibe"],
  ["windsurf", "Windsurf", ".windsurf"],
  ["zcode", "ZCode", ".zcode"],
].map(([id, name, root, detection = [root]]) =>
  Object.freeze({ id, name, root, detection: Object.freeze(detection) })));

const TOOL_BY_ID = new Map(TOOLS.map((tool) => [tool.id, tool]));
const SKILL_BY_ID = new Map(SKILLS.map((skill) => [skill.id, skill]));

export async function runOfficeKitCli(
  argv,
  {
    input = process.stdin,
    output = process.stdout,
    errorOutput = process.stderr,
  } = {},
) {
  const packageMetadata = JSON.parse(
    await readFile(path.join(PACKAGE_ROOT, "package.json"), "utf8"),
  );
  const [command, ...commandArguments] = argv;

  if (command === "run") {
    const { runTaskCommand } = await import("./run-task.mjs");
    await runTaskCommand(commandArguments, { output });
    return;
  }
  if (command === "repl") {
    const { runReplCommand } = await import("./repl.mjs");
    await runReplCommand(commandArguments, { input, output, errorOutput });
    return;
  }
  if (command === "template") {
    await runTemplateCommand(commandArguments, { output });
    return;
  }
  if (command === "excel") {
    const { runExcelCommand } = await import("../excel-live/cli.mjs");
    await runExcelCommand(commandArguments, { input, output });
    return;
  }
  if (command === "live") {
    const { runLiveCommand } = await import("../live/cli.mjs");
    await runLiveCommand(commandArguments, { input, output });
    return;
  }

  const parsed = parseArguments(argv);

  if (parsed.version) {
    output.write(`${packageMetadata.version}\n`);
    return;
  }
  if (parsed.help || parsed.command === "help") {
    output.write(helpText(packageMetadata.version));
    return;
  }
  if (!["init", "update"].includes(parsed.command)) {
    throw new Error(`Unknown command "${parsed.command}". Run "officekit --help".`);
  }

  const requestedPath = path.resolve(parsed.targetPath ?? ".");
  await mkdir(requestedPath, { recursive: true });
  const projectPath = await realpath(requestedPath);
  const projectStat = await lstat(projectPath);
  if (!projectStat.isDirectory()) {
    throw new Error(`Project path is not a directory: ${projectPath}`);
  }

  const manifest = await readManagedManifest(projectPath);
  if (parsed.command === "update" && manifest == null) {
    throw new Error(`OfficeKit is not initialized in ${projectPath}. Run "officekit init" first.`);
  }

  const toolIds = await selectTools({
    parsed,
    manifest,
    projectPath,
    input,
    output,
  });
  const result = await installSkills({
    command: parsed.command,
    force: parsed.force,
    manifest,
    packageMetadata,
    projectPath,
    toolIds,
  });

  if (parsed.json) {
    output.write(`${JSON.stringify({ ok: true, ...result })}\n`);
  } else {
    output.write(`OfficeKit ${parsed.command === "init" ? "initialized" : "updated"} in ${projectPath}\n`);
    for (const tool of result.tools) {
      output.write(`  ${tool.name}: ${tool.skillsRoot}\n`);
    }
    output.write(
      `  ${result.created} created, ${result.updated} updated, ${result.unchanged} unchanged\n\n` +
      "Start with /office-kit, or call a file-specific Skill directly.\n",
    );
  }

  if (errorOutput && result.warnings.length > 0) {
    for (const warning of result.warnings) errorOutput.write(`OfficeKit: ${warning}\n`);
  }
}

async function runTemplateCommand(args, { output }) {
  const [subcommand, ...subcommandArguments] = args;
  const {
    TEMPLATE_SEARCH_USAGE,
    formatTemplateSearchResult,
    parseTemplateSearchArguments,
    queryTemplates,
  } = await import("../templates/search.mjs");

  if (
    subcommand == null ||
    subcommand === "help" ||
    subcommand === "--help" ||
    subcommand === "-h"
  ) {
    output.write(`${TEMPLATE_SEARCH_USAGE}\n`);
    return;
  }
  if (subcommand !== "search") {
    throw new Error(
      `Unknown template command "${subcommand}". Run "officekit template search --help".`,
    );
  }

  const request = parseTemplateSearchArguments(subcommandArguments);
  if (request.help) {
    output.write(`${TEMPLATE_SEARCH_USAGE}\n`);
    return;
  }
  const { json = false, ...query } = request;
  const result = await queryTemplates({
    ...query,
    projectPath: process.cwd(),
  });
  output.write(
    json
      ? `${JSON.stringify(result)}\n`
      : `${formatTemplateSearchResult(result)}\n`,
  );
}

function parseArguments(argv) {
  const args = [...argv];
  if (args.length === 0) {
    return { command: "help", help: true };
  }

  let command;
  let targetPath;
  let tools;
  let force = false;
  let json = false;
  let yes = false;
  let help = false;
  let version = false;

  while (args.length > 0) {
    const argument = args.shift();
    if (argument === "--help" || argument === "-h") {
      help = true;
    } else if (argument === "--version" || argument === "-v") {
      version = true;
    } else if (argument === "--force") {
      force = true;
    } else if (argument === "--json") {
      json = true;
    } else if (argument === "--yes" || argument === "-y") {
      yes = true;
    } else if (argument === "--tools") {
      if (args.length === 0 || args[0].startsWith("-")) {
        throw new Error("--tools requires a comma-separated value or \"all\".");
      }
      tools = args.shift();
    } else if (argument.startsWith("--tools=")) {
      tools = argument.slice("--tools=".length);
    } else if (argument.startsWith("-")) {
      throw new Error(`Unknown option "${argument}". Run "officekit --help".`);
    } else if (command == null) {
      command = argument;
    } else if (targetPath == null) {
      targetPath = argument;
    } else {
      throw new Error(`Unexpected argument "${argument}".`);
    }
  }

  return {
    command: command ?? "help",
    targetPath,
    tools,
    force,
    json,
    yes,
    help,
    version,
  };
}

async function selectTools({ parsed, manifest, projectPath, input, output }) {
  const configured = manifest?.tools ?? [];
  let selected;

  if (parsed.tools != null) {
    selected = parseToolIds(parsed.tools);
  } else if (configured.length > 0) {
    selected = configured;
  } else {
    const detected = await detectTools(projectPath);
    const canPrompt = Boolean(input.isTTY && output.isTTY && !parsed.json && !parsed.yes);
    if (canPrompt) {
      selected = await promptForTools({ detected, input, output });
    } else if (detected.length > 0) {
      selected = detected;
    } else {
      throw new Error(
        "No Agent tool was detected. Run officekit init in an interactive terminal to choose a target, or pass --tools <ids>; for example, --tools codex or --tools claude,cursor.",
      );
    }
  }

  const combined = [...configured, ...selected];
  return [...new Set(combined)];
}

function parseToolIds(value) {
  const normalized = value.trim().toLowerCase();
  if (normalized === "all") return TOOLS.map((tool) => tool.id);
  if (normalized === "none" || normalized.length === 0) {
    throw new Error("OfficeKit initialization requires at least one Agent tool.");
  }
  const tokens = normalized.split(",").map((token) => token.trim()).filter(Boolean);
  if (tokens.includes("all") || tokens.includes("none")) {
    throw new Error("Do not combine \"all\" or \"none\" with specific tool IDs.");
  }
  const invalid = tokens.filter((token) => !TOOL_BY_ID.has(token));
  if (invalid.length > 0) {
    throw new Error(
      `Unknown Agent tool(s): ${invalid.join(", ")}. Available IDs: ${TOOLS.map((tool) => tool.id).join(", ")}`,
    );
  }
  return [...new Set(tokens)];
}

async function detectTools(projectPath) {
  const detected = [];
  for (const tool of TOOLS) {
    for (const marker of tool.detection) {
      if (await pathExists(path.join(projectPath, marker))) {
        detected.push(tool.id);
        break;
      }
    }
  }
  return detected;
}

async function promptForTools({ detected, input, output }) {
  const defaultIds = detected.length > 0 ? detected : ["agents"];
  output.write(
    detected.length > 0
      ? `Detected Agent tools: ${detected.join(", ")}\n`
      : "Choose where OfficeKit should install its eight project Skills.\n",
  );
  output.write(`Available: ${TOOLS.map((tool) => tool.id).join(", ")}\n`);
  const prompt = createInterface({ input, output });
  try {
    const answer = await prompt.question(`Agent tools [${defaultIds.join(",")}]: `);
    return answer.trim() === "" ? defaultIds : parseToolIds(answer);
  } finally {
    prompt.close();
  }
}

async function installSkills({
  command,
  force,
  manifest,
  packageMetadata,
  projectPath,
  toolIds,
}) {
  await assertSafeParents(projectPath, ".office-kit");
  const sourceHashes = new Map();
  for (const skill of SKILLS) {
    const sourcePath = path.join(PACKAGE_ROOT, skill.source);
    sourceHashes.set(skill.id, await hashDirectory(sourcePath));
  }

  const managedByPath = new Map(
    (manifest?.installations ?? []).map((installation) => [
      installation.path,
      installation,
    ]),
  );
  const plans = [];
  for (const toolId of toolIds) {
    const tool = TOOL_BY_ID.get(toolId);
    if (!tool) throw new Error(`Managed manifest refers to unsupported Agent tool "${toolId}".`);
    for (const skill of SKILLS) {
      const destination = toPosix(path.join(tool.root, "skills", skill.id));
      await assertSafeParents(projectPath, path.dirname(destination));
      const sourceHash = sourceHashes.get(skill.id);
      const managed = managedByPath.get(destination);
      const existing = await inspectDestination(projectPath, destination);
      let action = "create";
      if (existing != null) {
        if (managed == null) {
          throw new Error(
            `${destination} already exists and is not managed by OfficeKit. Move or remove it before initialization.`,
          );
        } else if (existing.sha256 === sourceHash) {
          action = "unchanged";
        } else if (existing.sha256 === managed.sha256 || force) {
          action = "update";
        } else {
          throw new Error(
            `${destination} was changed after OfficeKit installed it. Review the changes or rerun with --force.`,
          );
        }
      }
      plans.push({
        action,
        destination,
        source: path.join(PACKAGE_ROOT, skill.source),
        sourceHash,
        skill: skill.id,
        tool: tool.id,
      });
    }
  }

  for (const plan of plans) {
    if (plan.action !== "unchanged") {
      await replaceDirectoryAtomically({
        projectPath,
        relativeDestination: plan.destination,
        source: plan.source,
        sourceHash: plan.sourceHash,
      });
    }
    managedByPath.set(plan.destination, {
      tool: plan.tool,
      skill: plan.skill,
      path: plan.destination,
      sha256: plan.sourceHash,
    });
  }

  const installations = [...managedByPath.values()]
    .map(validateInstallation)
    .sort((left, right) =>
      left.path < right.path ? -1 : left.path > right.path ? 1 : 0);
  const nextManifest = {
    schemaVersion: 1,
    package: {
      name: packageMetadata.name,
      version: packageMetadata.version,
    },
    tools: [...new Set(installations.map((installation) => installation.tool))].sort(),
    skills: SKILLS.map((skill) => skill.id),
    installations,
  };
  await writeManagedManifest(projectPath, nextManifest);

  return {
    command,
    project: projectPath,
    package: nextManifest.package,
    manifest: path.join(projectPath, MANIFEST_PATH),
    tools: nextManifest.tools.map((toolId) => {
      const tool = TOOL_BY_ID.get(toolId);
      return {
        id: tool.id,
        name: tool.name,
        skillsRoot: path.join(projectPath, tool.root, "skills"),
      };
    }),
    skills: nextManifest.skills,
    created: plans.filter((plan) => plan.action === "create").length,
    updated: plans.filter((plan) => plan.action === "update").length,
    unchanged: plans.filter((plan) => plan.action === "unchanged").length,
    warnings: nextManifest.tools.includes("hermes")
      ? ["Hermes must be configured to load this project's .hermes/skills directory."]
      : [],
  };
}

async function inspectDestination(projectPath, relativeDestination) {
  const destination = path.join(projectPath, relativeDestination);
  let stat;
  try {
    stat = await lstat(destination);
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    throw error;
  }
  if (stat.isSymbolicLink() || !stat.isDirectory()) {
    throw new Error(`${relativeDestination} must be a regular directory, not a file or symbolic link.`);
  }
  return { sha256: await hashDirectory(destination) };
}

async function hashDirectory(root) {
  const rootStat = await lstat(root);
  if (rootStat.isSymbolicLink() || !rootStat.isDirectory()) {
    throw new Error(`Skill source must be a regular directory: ${root}`);
  }
  const hash = createHash("sha256");

  async function visit(directory, relativeDirectory) {
    const entries = await readdir(directory, { withFileTypes: true });
    entries.sort((left, right) =>
      left.name < right.name ? -1 : left.name > right.name ? 1 : 0);
    for (const entry of entries) {
      const absolute = path.join(directory, entry.name);
      const relative = toPosix(path.join(relativeDirectory, entry.name));
      const stat = await lstat(absolute);
      if (stat.isSymbolicLink()) {
        throw new Error(`Skill trees cannot contain symbolic links: ${absolute}`);
      }
      if (stat.isDirectory()) {
        hash.update(`D\0${relative}\0`);
        await visit(absolute, relative);
      } else if (stat.isFile()) {
        hash.update(`F\0${relative}\0${stat.size}\0`);
        hash.update(await readFile(absolute));
      } else {
        throw new Error(`Skill trees can contain only regular files and directories: ${absolute}`);
      }
    }
  }

  await visit(root, "");
  return hash.digest("hex");
}

async function replaceDirectoryAtomically({
  projectPath,
  relativeDestination,
  source,
  sourceHash,
}) {
  const destination = path.join(projectPath, relativeDestination);
  const parent = path.dirname(destination);
  await ensureSafeDirectory(projectPath, path.relative(projectPath, parent));
  const token = `${process.pid}-${randomUUID()}`;
  const staging = path.join(parent, `.officekit-${path.basename(destination)}-${token}.tmp`);
  const backup = path.join(parent, `.officekit-${path.basename(destination)}-${token}.bak`);
  let movedExisting = false;

  try {
    await cp(source, staging, {
      recursive: true,
      force: false,
      errorOnExist: true,
      preserveTimestamps: false,
    });
    const stagedHash = await hashDirectory(staging);
    if (stagedHash !== sourceHash) {
      throw new Error(`Staged Skill verification failed for ${relativeDestination}.`);
    }
    if (await pathExists(destination)) {
      await rename(destination, backup);
      movedExisting = true;
    }
    await rename(staging, destination);
    if (movedExisting) await rm(backup, { recursive: true, force: true });
  } catch (error) {
    await rm(staging, { recursive: true, force: true });
    if (movedExisting && !(await pathExists(destination)) && await pathExists(backup)) {
      await rename(backup, destination);
    }
    throw error;
  }
}

async function readManagedManifest(projectPath) {
  const manifestFile = path.join(projectPath, MANIFEST_PATH);
  const stat = await lstatIfExists(manifestFile);
  if (stat == null) return null;
  await assertSafeParents(projectPath, path.dirname(MANIFEST_PATH));
  if (stat.isSymbolicLink() || !stat.isFile()) {
    throw new Error(`${MANIFEST_PATH} must be a regular file.`);
  }
  let manifest;
  try {
    manifest = JSON.parse(await readFile(manifestFile, "utf8"));
  } catch (error) {
    throw new Error(`Cannot read ${MANIFEST_PATH}: ${error.message}`);
  }
  if (manifest?.schemaVersion !== 1 || !Array.isArray(manifest.tools) ||
      !Array.isArray(manifest.installations)) {
    throw new Error(`${MANIFEST_PATH} uses an unsupported or invalid schema.`);
  }
  for (const toolId of manifest.tools) {
    if (!TOOL_BY_ID.has(toolId)) {
      throw new Error(`${MANIFEST_PATH} refers to unsupported Agent tool "${toolId}".`);
    }
  }
  manifest.installations = manifest.installations.map(validateInstallation);
  return manifest;
}

function validateInstallation(installation) {
  const tool = TOOL_BY_ID.get(installation?.tool);
  const skill = SKILL_BY_ID.get(installation?.skill);
  if (!tool || !skill || typeof installation?.path !== "string" ||
      !/^[a-f0-9]{64}$/.test(installation?.sha256 ?? "")) {
    throw new Error(`${MANIFEST_PATH} contains an invalid installation record.`);
  }
  const expected = toPosix(path.join(tool.root, "skills", skill.id));
  if (installation.path !== expected) {
    throw new Error(`${MANIFEST_PATH} contains an out-of-scope installation path.`);
  }
  return {
    tool: tool.id,
    skill: skill.id,
    path: expected,
    sha256: installation.sha256,
  };
}

async function writeManagedManifest(projectPath, manifest) {
  await ensureSafeDirectory(projectPath, ".office-kit");
  const destination = path.join(projectPath, MANIFEST_PATH);
  const existing = await lstatIfExists(destination);
  if (existing != null) {
    const stat = existing;
    if (stat.isSymbolicLink() || !stat.isFile()) {
      throw new Error(`${MANIFEST_PATH} must be a regular file.`);
    }
  }
  const temporary = `${destination}.${process.pid}-${randomUUID()}.tmp`;
  const backup = `${destination}.${process.pid}-${randomUUID()}.bak`;
  await writeFile(temporary, `${JSON.stringify(manifest, null, 2)}\n`, {
    encoding: "utf8",
    flag: "wx",
    mode: 0o600,
  });
  let movedExisting = false;
  try {
    if (existing != null) {
      await rename(destination, backup);
      movedExisting = true;
    }
    await rename(temporary, destination);
    if (movedExisting) await rm(backup, { force: true });
  } catch (error) {
    await rm(temporary, { force: true });
    if (movedExisting && !(await pathExists(destination)) && await pathExists(backup)) {
      await rename(backup, destination);
    }
    throw error;
  }
}

async function assertSafeParents(projectPath, relativePath) {
  const normalized = path.normalize(relativePath);
  if (path.isAbsolute(normalized) || normalized.split(path.sep).includes("..")) {
    throw new Error(`Path escapes the project: ${relativePath}`);
  }
  let current = projectPath;
  for (const segment of normalized.split(path.sep).filter(Boolean)) {
    current = path.join(current, segment);
    let stat;
    try {
      stat = await lstat(current);
    } catch (error) {
      if (error?.code === "ENOENT") return;
      throw error;
    }
    if (stat.isSymbolicLink() || !stat.isDirectory()) {
      throw new Error(`${path.relative(projectPath, current)} must be a regular directory.`);
    }
  }
}

async function ensureSafeDirectory(projectPath, relativePath) {
  const normalized = path.normalize(relativePath);
  if (path.isAbsolute(normalized) || normalized.split(path.sep).includes("..")) {
    throw new Error(`Path escapes the project: ${relativePath}`);
  }
  let current = projectPath;
  for (const segment of normalized.split(path.sep).filter(Boolean)) {
    current = path.join(current, segment);
    try {
      const stat = await lstat(current);
      if (stat.isSymbolicLink() || !stat.isDirectory()) {
        throw new Error(`${path.relative(projectPath, current)} must be a regular directory.`);
      }
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
      await mkdir(current);
    }
  }
}

async function pathExists(target) {
  return access(target).then(() => true, () => false);
}

async function lstatIfExists(target) {
  try {
    return await lstat(target);
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    throw error;
  }
}

function toPosix(value) {
  return value.split(path.sep).join("/");
}

function helpText(version) {
  return `OfficeKit ${version}

Install Skills, run Office tasks, and search reusable templates.

Usage:
  officekit init [path] [--tools <ids>] [--yes] [--json]
  officekit update [path] [--tools <ids>] [--force] [--json]
  officekit run <task.mjs> [-- <task arguments>]
  officekit repl [options]
  officekit template search [search options] [--json]
  officekit excel <command> [options]
  officekit live <command> --app <excel|powerpoint> [options]
  officekit --version

Commands:
  init       Choose Agent targets and install the OfficeKit Skills
  update     Refresh Skills already managed by OfficeKit
  run        Run a task with this OfficeKit installation
  repl       Run a persistent JSONL JavaScript task session
  template   Search the bundled and project template catalogs
  excel      Connect an open Microsoft Excel workbook to local OfficeKit control
  live       Connect a supported open Office document to local OfficeKit control

Options:
  --tools <ids>  Comma-separated Agent tool IDs, or "all"
  --yes, -y      Use detected tools without an interactive prompt
  --force        Replace locally changed OfficeKit-managed Skill trees
  --json         Print one machine-readable result
  --help, -h     Show this help
  --version, -v  Show the OfficeKit version

Common tool IDs:
  agents, claude, cursor, github-copilot, gemini, opencode

Examples:
  officekit init
  officekit init --tools claude,cursor
  officekit update
  officekit run task.mjs -- input.docx output.docx
  officekit template search --kind presentation --purpose "quarterly business review"
  officekit excel install
  officekit excel sessions --json
  officekit live install --app powerpoint
  officekit live sessions --app powerpoint --json
`;
}
