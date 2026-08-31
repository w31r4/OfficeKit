import { lstat, readFile, realpath } from "node:fs/promises";
import path from "node:path";

import { inspectImageBytes } from "../shared/image-bytes.mjs";
import { auditPresentationImages, writeImageSourcesSidecar } from "./audit.mjs";
import { downloadRemoteImage } from "./download.mjs";
import { imageError } from "./errors.mjs";
import { materializeLucideIcon } from "./lucide.mjs";
import { searchImageCandidates } from "./providers.mjs";
import {
  MAX_IMAGE_BYTES,
  MAX_IMAGE_DIMENSION,
  MAX_IMAGE_PIXELS,
  addTaskImageAsset,
  imageTaskState,
  openImageTask,
  recordTaskImageSearch,
  resolveTaskImageCandidate,
} from "./task-assets.mjs";

const REMOTE_DECLARED_RIGHTS = new Set(["permission", "public-domain", "cc0", "cc-by", "official-press-kit"]);

export const IMAGE_USAGE = `Usage:
  officekit image search "<query>" --task <task-id> --kind <photo|illustration|icon> --purpose <hero|evidence|context|decoration> --orientation <landscape|portrait|square> [--max <1..20>] [--json]
  officekit image add --task <task-id> --candidate <candidate-ref> [--json]
  officekit image add --task <task-id> --file <path> --rights <user-provided|generated|permission|public-domain|cc0|cc-by> [rights options] [--json]
  officekit image add --task <task-id> --url <https-url> --source-page <https-url> --rights <permission|public-domain|cc0|cc-by|official-press-kit> [rights options] [--json]
  officekit image list --task <task-id> [--json]
  officekit image audit <presentation.pptx> --task <task-id> [--sources-output <path>] [--json]

Common options:
  --workspace <path>    Workspace containing .office-kit/tasks (default: current directory)
  --author <name>       Required for CC BY
  --license-url <url>   Required for CC BY
  --json                Print one machine-readable result

Search returns candidates only; selectionMade is always false.`;

function parse(args) {
  const [command, ...rest] = args;
  if (!command || ["help", "--help", "-h"].includes(command)) return { command: "help" };
  const parsed = { command, positionals: [], json: false };
  const valueOptions = new Map([
    ["--task", "taskId"],
    ["--workspace", "workspaceRoot"],
    ["--kind", "kind"],
    ["--purpose", "purpose"],
    ["--orientation", "orientation"],
    ["--max", "max"],
    ["--candidate", "candidateRef"],
    ["--file", "file"],
    ["--url", "url"],
    ["--source-page", "sourcePage"],
    ["--rights", "rights"],
    ["--author", "author"],
    ["--license-url", "licenseUrl"],
    ["--sources-output", "sourcesOutput"],
  ]);
  for (let index = 0; index < rest.length; index += 1) {
    const argument = rest[index];
    if (argument === "--json") parsed.json = true;
    else if (argument === "--help" || argument === "-h") parsed.help = true;
    else if (valueOptions.has(argument)) {
      const value = rest[++index];
      if (value == null || value.startsWith("--")) throw imageError("invalid-image-command", `${argument} requires a value.`);
      parsed[valueOptions.get(argument)] = value;
    } else if (argument.startsWith("--")) {
      throw imageError("invalid-image-command", `Unknown image option ${argument}.`);
    } else {
      parsed.positionals.push(argument);
    }
  }
  return parsed;
}

function required(value, label) {
  if (value == null || String(value).trim() === "") throw imageError("invalid-image-command", `${label} is required.`);
  return String(value);
}

function rightsMetadata(parsed, extras = {}) {
  return {
    author: parsed.author,
    licenseUrl: parsed.licenseUrl,
    sourcePage: parsed.sourcePage,
    evidence: "user-declared",
    ...extras,
  };
}

async function localImage(parsed) {
  const requested = path.resolve(required(parsed.file, "--file"));
  const stat = await lstat(requested).catch((error) => {
    if (error?.code === "ENOENT") throw imageError("image-file-missing", `Image file does not exist: ${requested}`);
    throw error;
  });
  if (stat.isSymbolicLink() || !stat.isFile()) throw imageError("unsafe-image-path", "Image input must be a regular non-symlink file.");
  if (stat.size > MAX_IMAGE_BYTES) throw imageError("image-file-too-large", `Image input exceeds ${MAX_IMAGE_BYTES} bytes.`);
  const canonical = await realpath(requested);
  const bytes = await readFile(canonical);
  const inspected = inspectImageBytes(bytes, {
    label: "Local image",
    maxBytes: MAX_IMAGE_BYTES,
    maxPixels: MAX_IMAGE_PIXELS,
    maxDimension: MAX_IMAGE_DIMENSION,
  });
  return {
    bytes,
    mimeType: inspected.mimeType,
    rights: required(parsed.rights, "--rights"),
    rightsMetadata: rightsMetadata(parsed),
    source: { kind: "file", originalPath: canonical, fileName: path.basename(canonical) },
  };
}

async function candidateImage(task, parsed, downloader) {
  const candidateRef = required(parsed.candidateRef, "--candidate");
  const candidate = await resolveTaskImageCandidate(task, candidateRef);
  if (String(candidate.acquisitionUrl || "").startsWith("lucide:")) {
    const materialized = await materializeLucideIcon(candidate.acquisitionUrl);
    return {
      bytes: materialized.bytes,
      mimeType: materialized.mimeType,
      rights: candidate.rights.rights,
      rightsMetadata: candidate.rights,
      source: { ...materialized.source, kind: "candidate", candidateRef, provider: candidate.provider },
    };
  }
  const downloaded = await downloader(candidate.acquisitionUrl);
  return {
    bytes: downloaded.bytes,
    mimeType: downloaded.mimeType,
    rights: candidate.rights.rights,
    rightsMetadata: candidate.rights,
    source: {
      kind: "candidate",
      candidateRef,
      provider: candidate.provider,
      sourcePage: candidate.sourcePage,
      requestedUrl: candidate.acquisitionUrl,
      finalUrl: downloaded.finalUrl,
      redirects: downloaded.redirects,
    },
  };
}

async function remoteImage(parsed, downloader) {
  const rights = required(parsed.rights, "--rights").toLowerCase();
  if (!REMOTE_DECLARED_RIGHTS.has(rights)) {
    throw imageError("image-rights-blocked", `Remote URL rights ${rights} are not allowed.`);
  }
  const sourcePage = required(parsed.sourcePage, "--source-page");
  const requestedUrl = required(parsed.url, "--url");
  const downloaded = await downloader(requestedUrl);
  return {
    bytes: downloaded.bytes,
    mimeType: downloaded.mimeType,
    rights,
    rightsMetadata: rightsMetadata(parsed, { sourcePage }),
    source: { kind: "url", sourcePage, requestedUrl, finalUrl: downloaded.finalUrl, redirects: downloaded.redirects },
  };
}

function exactlyOneSource(parsed) {
  const values = [parsed.candidateRef, parsed.file, parsed.url].filter((value) => value != null);
  if (values.length !== 1) throw imageError("invalid-image-command", "image add requires exactly one of --candidate, --file, or --url.");
}

function humanSearch(result) {
  if (result.candidates.length === 0) return `No compliant image candidates found for “${result.query}”.\nSelection remains with the Agent.`;
  const rows = result.candidates.map((candidate, index) => `${index + 1}. ${candidate.candidateRef}  ${candidate.provider}  ${candidate.title || "Untitled"}  ${candidate.rights}`);
  return [`Image candidates for “${result.query}”`, ...rows, "Selection remains with the Agent."].join("\n");
}

function humanList(result) {
  const lines = [`Task ${result.taskId}: ${result.assets.length} image assets, ${result.searches.length} searches`];
  for (const asset of result.assets) lines.push(`${asset.sha256.slice(0, 12)}  ${asset.mimeType}  ${asset.width}x${asset.height}  ${asset.rights}  ${asset.path}`);
  return lines.join("\n");
}

export async function runImageCommand(args, {
  output = process.stdout,
  searcher = searchImageCandidates,
  downloader = downloadRemoteImage,
} = {}) {
  const parsed = parse(args);
  if (parsed.command === "help" || parsed.help) {
    output.write(`${IMAGE_USAGE}\n`);
    return;
  }
  if (!new Set(["search", "add", "list", "audit"]).has(parsed.command)) {
    throw imageError("invalid-image-command", `Unknown image command ${parsed.command}.`);
  }
  const taskId = required(parsed.taskId, "--task");
  const workspaceRoot = path.resolve(parsed.workspaceRoot || process.cwd());
  const task = await openImageTask({ workspaceRoot, taskId });
  let result;

  if (parsed.command === "search") {
    if (parsed.positionals.length !== 1) throw imageError("invalid-image-command", "image search requires one quoted query.");
    const found = await searcher({
      query: parsed.positionals[0],
      kind: required(parsed.kind, "--kind"),
      purpose: required(parsed.purpose, "--purpose"),
      orientation: required(parsed.orientation, "--orientation"),
      max: parsed.max,
    });
    const recorded = await recordTaskImageSearch(task, found);
    result = { ok: true, command: "search", ...recorded };
    output.write(parsed.json ? `${JSON.stringify(result)}\n` : `${humanSearch(result)}\n`);
    return result;
  }

  if (parsed.command === "add") {
    if (parsed.positionals.length !== 0) throw imageError("invalid-image-command", "image add does not accept positional arguments.");
    exactlyOneSource(parsed);
    const input = parsed.candidateRef
      ? await candidateImage(task, parsed, downloader)
      : parsed.file
        ? await localImage(parsed)
        : await remoteImage(parsed, downloader);
    const asset = await addTaskImageAsset(task, input);
    result = { ok: true, command: "add", taskId, asset };
    output.write(parsed.json ? `${JSON.stringify(result)}\n` : `${asset.path}\n`);
    return result;
  }

  if (parsed.command === "list") {
    if (parsed.positionals.length !== 0) throw imageError("invalid-image-command", "image list does not accept positional arguments.");
    const state = await imageTaskState({ workspaceRoot, taskId });
    result = { ok: true, command: "list", taskId, assets: state.assets, searches: state.searches };
    output.write(parsed.json ? `${JSON.stringify(result)}\n` : `${humanList(result)}\n`);
    return result;
  }

  if (parsed.positionals.length !== 1) throw imageError("invalid-image-command", "image audit requires one presentation.pptx path.");
  const audit = await auditPresentationImages(task, { pptxPath: parsed.positionals[0] });
  const sources = parsed.sourcesOutput ? await writeImageSourcesSidecar(audit, parsed.sourcesOutput) : undefined;
  result = { ok: true, command: "audit", taskId, audit, ...(sources ? { sources } : {}) };
  output.write(parsed.json ? `${JSON.stringify(result)}\n` : `${audit.presentation.path}\n${audit.used.length} registered used; ${audit.unregistered.length} unregistered media; ${audit.unused.length} unused assets\n`);
  return result;
}
