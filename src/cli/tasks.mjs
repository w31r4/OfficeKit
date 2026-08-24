import path from "node:path";

import { deleteTask, listTasks, taskDetail } from "./task-store.mjs";

export const TASKS_USAGE = [
  "Usage:",
  "  officekit tasks [--all] [--json] [--workspace <path>]",
  "  officekit tasks <task-id> [--json] [--workspace <path>]",
  "  officekit tasks --delete <task-id> --yes [--json] [--workspace <path>]",
].join("\n");

export async function runTasksCommand(args, { output = process.stdout } = {}) {
  const options = parseTasksArguments(args);
  if (options.help) {
    output.write(`${TASKS_USAGE}\n`);
    return;
  }
  if (options.deleteTaskId) {
    if (!options.yes) throw new Error("Deleting a task requires --yes.");
    const result = await deleteTask({ workspaceRoot: options.workspaceRoot, taskId: options.deleteTaskId });
    output.write(options.json
      ? `${JSON.stringify({ ok: true, ...result })}\n`
      : `Deleted OfficeKit task ${result.taskId} (${formatBytes(result.bytes)}) from ${result.workspace}\n`);
    return;
  }
  if (options.taskId) {
    const result = await taskDetail({ workspaceRoot: options.workspaceRoot, taskId: options.taskId });
    output.write(options.json ? `${JSON.stringify(result)}\n` : `${formatTaskDetail(result)}\n`);
    return;
  }
  const result = await listTasks({ workspaceRoot: options.workspaceRoot, all: options.all });
  output.write(options.json ? `${JSON.stringify(result)}\n` : `${formatTaskList(result)}\n`);
}

export function parseTasksArguments(args) {
  const values = [...args];
  const options = { all: false, json: false, yes: false, help: false, workspaceRoot: undefined, taskId: undefined, deleteTaskId: undefined };
  while (values.length > 0) {
    const value = values.shift();
    if (value === "--all") options.all = true;
    else if (value === "--json") options.json = true;
    else if (value === "--yes" || value === "-y") options.yes = true;
    else if (value === "--help" || value === "-h") options.help = true;
    else if (value === "--workspace") options.workspaceRoot = required(values, value);
    else if (value.startsWith("--workspace=")) options.workspaceRoot = value.slice(12);
    else if (value === "--delete") options.deleteTaskId = required(values, value);
    else if (value.startsWith("--delete=")) options.deleteTaskId = value.slice(9);
    else if (value.startsWith("-")) throw new Error(`Unknown tasks option: ${value}.`);
    else if (options.taskId == null) options.taskId = value;
    else throw new Error(`Unexpected tasks argument: ${value}.`);
  }
  if (options.taskId && options.deleteTaskId) throw new Error("Choose task detail or --delete, not both.");
  if (options.taskId && options.all) throw new Error("--all cannot be combined with one task ID.");
  return options;
}

export function formatTaskList(result) {
  const lines = [`OfficeKit · ${result.workspace} · ${result.total} task${result.total === 1 ? "" : "s"}`];
  if (result.tasks.length === 0) {
    lines.push("", "No OfficeKit tasks in this workspace.");
    return lines.join("\n");
  }
  const idWidth = Math.max(2, ...result.tasks.map((task) => task.id.length));
  const goalWidth = Math.min(36, Math.max(4, ...result.tasks.map((task) => displayWidth(task.goal))));
  lines.push("", `${pad("ID", idWidth)}  ${pad("GOAL", goalWidth)}  ${pad("HEAD", 6)}  ${pad("STATE", 9)}  UPDATED`);
  for (const task of result.tasks) {
    lines.push(`${pad(task.id, idWidth)}  ${pad(truncate(task.goal, goalWidth), goalWidth)}  ${pad(task.head?.id || "—", 6)}  ${pad(task.state, 9)}  ${relativeTime(task.updatedAt)}`);
  }
  if (result.truncated) lines.push("", `${result.total - result.shown} more task${result.total - result.shown === 1 ? "" : "s"}. Use officekit tasks --all.`);
  if (result.invalid.length > 0) lines.push("", `${result.invalid.length} invalid task entr${result.invalid.length === 1 ? "y" : "ies"} ignored.`);
  return lines.join("\n");
}

export function formatTaskDetail(result) {
  const task = result.task;
  const lines = [
    `Task: ${task.id}`,
    `Workspace: ${result.workspace}`,
    `Goal: ${task.goal}`,
    `State: ${task.state}`,
  ];
  if (task.inputs.length) {
    lines.push("", "Inputs");
    for (const input of task.inputs) lines.push(`  ${input.name} · ${input.kind} · ${input.sha256.slice(0, 12)}`);
  }
  if (task.artifacts.length) {
    lines.push("", "Artifacts");
    for (const artifact of task.artifacts) lines.push(`  ${artifact.name} · ${artifact.kind}${artifact.headRevision ? ` · ${artifact.headRevision.sha256.slice(0, 12)}` : " · no commit"}`);
  }
  if (task.plan) {
    lines.push("", "Plan");
    lines.push(`  ${task.plan.mode} · ${task.plan.pageCount} page${task.plan.pageCount === 1 ? "" : "s"} · ${task.plan.recipe}`);
    lines.push(`  ${task.plan.state} · ${task.plan.sha256.slice(0, 12)} · ${formatBytes(task.plan.bytes)}`);
  }
  lines.push("", "Head");
  lines.push(task.head
    ? `  ${task.head.id} · ${task.head.reviewVerdict} · visual ${task.head.visualReview}\n  ${task.head.summary}`
    : "  none · starts from staged inputs");
  if (task.pending.length) {
    lines.push("", "Attention");
    for (const pending of task.pending.slice(0, 5)) lines.push(`  ${pending.summary || pending.type}${pending.maybeApplied ? " · maybe applied" : ""}`);
    if (task.pending.length > 5) lines.push(`  ${task.pending.length - 5} more item(s)`);
  }
  if (task.next) lines.push("", "Next", `  ${task.next}`);
  if (task.publication) lines.push("", "Output", `  ${task.publication.path}`);
  lines.push("", `Storage: ${formatBytes(task.storageBytes)} · updated ${relativeTime(task.updatedAt)}`);
  return lines.join("\n");
}

function required(values, option) {
  const value = values.shift();
  if (value == null || value.startsWith("-")) throw new Error(`${option} requires a value.`);
  return value;
}

function truncate(value, maximum) {
  const text = String(value);
  if (displayWidth(text) <= maximum) return text;
  let result = "";
  let width = 0;
  for (const character of text) {
    const next = characterWidth(character);
    if (width + next + 1 > maximum) break;
    result += character;
    width += next;
  }
  return `${result}…`;
}

function displayWidth(value) {
  return [...String(value)].reduce((total, character) => total + characterWidth(character), 0);
}

function characterWidth(character) {
  const codePoint = character.codePointAt(0);
  return codePoint >= 0x1100 && (
    codePoint <= 0x115f ||
    codePoint === 0x2329 || codePoint === 0x232a ||
    (codePoint >= 0x2e80 && codePoint <= 0xa4cf && codePoint !== 0x303f) ||
    (codePoint >= 0xac00 && codePoint <= 0xd7a3) ||
    (codePoint >= 0xf900 && codePoint <= 0xfaff) ||
    (codePoint >= 0xfe10 && codePoint <= 0xfe19) ||
    (codePoint >= 0xfe30 && codePoint <= 0xfe6f) ||
    (codePoint >= 0xff00 && codePoint <= 0xff60) ||
    (codePoint >= 0xffe0 && codePoint <= 0xffe6) ||
    (codePoint >= 0x1f300 && codePoint <= 0x1faff)
  ) ? 2 : 1;
}

function pad(value, width) {
  const text = String(value);
  return `${text}${" ".repeat(Math.max(0, width - displayWidth(text)))}`;
}

function relativeTime(value, now = Date.now()) {
  const delta = Math.max(0, now - Date.parse(value));
  if (delta < 60_000) return "now";
  if (delta < 3_600_000) return `${Math.floor(delta / 60_000)}m`;
  if (delta < 86_400_000) return `${Math.floor(delta / 3_600_000)}h`;
  return `${Math.floor(delta / 86_400_000)}d`;
}

function formatBytes(value) {
  const bytes = Number(value) || 0;
  if (bytes < 1_024) return `${bytes} B`;
  if (bytes < 1_048_576) return `${(bytes / 1_024).toFixed(1)} KiB`;
  return `${(bytes / 1_048_576).toFixed(1)} MiB`;
}
