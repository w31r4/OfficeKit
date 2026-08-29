import { lstat, mkdir, readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { projectPptxToPpj } from "./native.mjs";
import { renderPpj, reviewPpj } from "./render-review.mjs";
import {
  resolveRegularFile,
  compilePpjWorkspace,
  loadPpjWorkspace,
  prettyProgram,
  replaceRegularFile,
  sha256,
  validatePpjWorkspace,
  writeExclusiveFile,
  writeImmutableContent,
} from "./workspace.mjs";

export const PPJ_USAGE = `Usage:
  officekit ppj import <input.pptx> -o <deck.ppj> [--json]
  officekit ppj inspect <deck.ppj> [--query <text>] [--page <id>] [--json]
  officekit ppj check <deck.ppj> [--fix] [--json]
  officekit ppj build <deck.ppj> -o <deck.pptx> [--json]
  officekit ppj render <deck.ppj> -o <previews/> [--pages <spec>] [--json]
  officekit ppj review <deck.ppj> [--json]`;

export async function runPpjCommand(args, {
  output = process.stdout,
  cwd = process.cwd(),
} = {}) {
  const [subcommand, ...rest] = args;
  if (subcommand == null || ["help", "--help", "-h"].includes(subcommand)) {
    output.write(`${PPJ_USAGE}\n`);
    return;
  }
  const parser = {
    import: parseImportArguments,
    inspect: parseInspectArguments,
    check: parseCheckArguments,
    build: parseBuildArguments,
    render: parseRenderArguments,
    review: parseReviewArguments,
  }[subcommand];
  if (!parser) throw new Error(`Unknown or unavailable PPJ command "${subcommand}". Run "officekit ppj --help".`);
  const request = parser(rest);
  if (request.help) {
    output.write(`${PPJ_USAGE}\n`);
    return;
  }
  const handler = {
    import: importPptxAsPpj,
    inspect: inspectPpj,
    check: checkPpj,
    build: buildPpj,
    render: renderPpj,
    review: reviewPpj,
  }[subcommand];
  const result = await handler(request, { cwd });
  output.write(request.json ? `${JSON.stringify(result)}\n` : `${formatResult(result)}\n`);
}

function parseImportArguments(args) {
  const positional = [];
  let outputPath;
  let json = false;
  let help = false;
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--json") json = true;
    else if (argument === "--help" || argument === "-h") help = true;
    else if (argument === "-o" || argument === "--output") {
      outputPath = requiredValue(args, ++index, argument);
    } else if (argument.startsWith("--output=")) outputPath = argument.slice("--output=".length);
    else if (argument.startsWith("-")) throw new Error(`Unknown PPJ import option "${argument}".`);
    else positional.push(argument);
  }
  if (help) return { help, json };
  if (positional.length !== 1 || !outputPath) {
    throw new Error("PPJ import requires one input.pptx and -o <deck.ppj>.");
  }
  return { inputPath: positional[0], outputPath, json, help };
}

function parseInspectArguments(args) {
  const positional = [];
  let query;
  let page;
  let json = false;
  let help = false;
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--json") json = true;
    else if (argument === "--help" || argument === "-h") help = true;
    else if (argument === "--query") query = requiredValue(args, ++index, argument);
    else if (argument.startsWith("--query=")) query = argument.slice("--query=".length);
    else if (argument === "--page") page = requiredValue(args, ++index, argument);
    else if (argument.startsWith("--page=")) page = argument.slice("--page=".length);
    else if (argument.startsWith("-")) throw new Error(`Unknown PPJ inspect option "${argument}".`);
    else positional.push(argument);
  }
  if (help) return { help, json };
  if (positional.length !== 1) throw new Error("PPJ inspect requires one deck.ppj input.");
  if (query != null && !query.trim()) throw new Error("PPJ inspect --query must not be empty.");
  return { inputPath: positional[0], query, page, json, help };
}

function parseCheckArguments(args) {
  const positional = [];
  let fix = false;
  let json = false;
  let help = false;
  for (const argument of args) {
    if (argument === "--json") json = true;
    else if (argument === "--fix") fix = true;
    else if (argument === "--help" || argument === "-h") help = true;
    else if (argument.startsWith("-")) throw new Error(`Unknown PPJ check option "${argument}".`);
    else positional.push(argument);
  }
  if (help) return { help, json };
  if (positional.length !== 1) throw new Error("PPJ check requires one deck.ppj input.");
  return { inputPath: positional[0], fix, json, help };
}

function parseBuildArguments(args) {
  const positional = [];
  let outputPath;
  let json = false;
  let help = false;
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--json") json = true;
    else if (argument === "--help" || argument === "-h") help = true;
    else if (argument === "-o" || argument === "--output") outputPath = requiredValue(args, ++index, argument);
    else if (argument.startsWith("--output=")) outputPath = argument.slice("--output=".length);
    else if (argument.startsWith("-")) throw new Error(`Unknown PPJ build option "${argument}".`);
    else positional.push(argument);
  }
  if (help) return { help, json };
  if (positional.length !== 1 || !outputPath) throw new Error("PPJ build requires one deck.ppj and -o <deck.pptx>.");
  return { inputPath: positional[0], outputPath, json, help };
}

function parseRenderArguments(args) {
  const positional = [];
  let outputPath;
  let pages;
  let json = false;
  let help = false;
  for (let index = 0; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--json") json = true;
    else if (argument === "--help" || argument === "-h") help = true;
    else if (argument === "-o" || argument === "--output") outputPath = requiredValue(args, ++index, argument);
    else if (argument.startsWith("--output=")) outputPath = argument.slice("--output=".length);
    else if (argument === "--pages") pages = requiredValue(args, ++index, argument);
    else if (argument.startsWith("--pages=")) pages = argument.slice("--pages=".length);
    else if (argument.startsWith("-")) throw new Error(`Unknown PPJ render option "${argument}".`);
    else positional.push(argument);
  }
  if (help) return { help, json };
  if (positional.length !== 1 || !outputPath) throw new Error("PPJ render requires one deck.ppj and -o <previews/>.");
  return { inputPath: positional[0], outputPath, pages, json, help };
}

function parseReviewArguments(args) {
  const positional = [];
  let json = false;
  let help = false;
  for (const argument of args) {
    if (argument === "--json") json = true;
    else if (argument === "--help" || argument === "-h") help = true;
    else if (argument.startsWith("-")) throw new Error(`Unknown PPJ review option "${argument}".`);
    else positional.push(argument);
  }
  if (help) return { help, json };
  if (positional.length !== 1) throw new Error("PPJ review requires one deck.ppj input.");
  return { inputPath: positional[0], json, help };
}

function requiredValue(args, index, option) {
  const value = args[index];
  if (!value || value.startsWith("-")) throw new Error(`${option} requires a value.`);
  return value;
}

export async function importPptxAsPpj(
  { inputPath, outputPath },
  { cwd = process.cwd(), project = projectPptxToPpj } = {},
) {
  const input = await resolveRegularFile(path.resolve(cwd, inputPath), "PPTX input");
  const destination = path.resolve(cwd, outputPath);
  if (path.extname(input).toLowerCase() !== ".pptx") throw new Error(`PPJ import input must be a .pptx file: ${input}`);
  if (path.extname(destination).toLowerCase() !== ".ppj") throw new Error(`PPJ import output must be a .ppj file: ${destination}`);
  if (input === destination) throw new Error("PPJ import cannot overwrite its PPTX input.");

  const sourceBytes = await readFile(input);
  const sourceSha256 = sha256(sourceBytes);
  const stem = path.basename(destination, ".ppj");
  const assetDirectoryName = `${stem}.assets`;
  const sourceRelative = `${assetDirectoryName}/source/${sourceSha256}.pptx`;
  const mediaRelative = `${assetDirectoryName}/media`;
  const projected = await project(sourceBytes, {
    sourceUri: sourceRelative,
    assetRootUri: mediaRelative,
    includeNodeMap: true,
  });
  if (!projected.sourceBound || projected.sourceSha256 !== sourceSha256) {
    throw new Error("OfficeKit projection did not bind the exact PPTX source revision.");
  }
  if (!/^[a-f0-9]{64}$/u.test(projected.programSha256) || sha256(projected.programJson) !== projected.programSha256) {
    throw new Error("OfficeKit projection returned a PPJ revision with an invalid content hash.");
  }

  const root = path.dirname(destination);
  await mkdir(root, { recursive: true });
  const sourceTarget = path.join(root, ...sourceRelative.split("/"));
  await writeImmutableContent(sourceTarget, sourceBytes, sourceSha256);
  const assets = [];
  for (const asset of projected.assets) {
    if (!/^[a-f0-9]{64}$/u.test(asset.sha256) || sha256(asset.data) !== asset.sha256) {
      throw new Error(`Projected PPJ asset ${asset.id} failed its content hash.`);
    }
    if (!asset.fileName || path.basename(asset.fileName) !== asset.fileName) {
      throw new Error(`Projected PPJ asset ${asset.id} has an unsafe file name.`);
    }
    const relative = `${mediaRelative}/${asset.fileName}`;
    const target = path.join(root, ...relative.split("/"));
    await writeImmutableContent(target, asset.data, asset.sha256);
    assets.push(Object.freeze({
      id: asset.id,
      path: target,
      uri: relative,
      mimeType: asset.mimeType,
      sha256: asset.sha256,
    }));
  }
  try {
    await writeExclusiveFile(destination, projected.programJson, 0o644);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(`PPJ output already exists: ${destination}`);
    throw error;
  }
  return Object.freeze({
    ok: true,
    command: "import",
    input,
    output: destination,
    programSha256: projected.programSha256,
    source: Object.freeze({ path: sourceTarget, uri: sourceRelative, sha256: sourceSha256 }),
    sourceBound: true,
    pages: programPageCount(projected.programJson),
    expandedElementCount: projected.expandedElementCount,
    assets: Object.freeze(assets),
    diagnostics: projected.diagnostics,
  });
}

export async function inspectPpj(
  { inputPath, query, page },
  { cwd = process.cwd(), load = loadPpjWorkspace, validate = validatePpjWorkspace } = {},
) {
  const workspace = await load(inputPath, { cwd });
  const validated = await validate(workspace, { includeNodeMap: true });
  const program = JSON.parse(Buffer.from(validated.programJson).toString("utf8"));
  const indexed = indexProgram(program);
  if (page != null && !indexed.pages.has(page)) throw new Error(`PPJ page ID does not exist: ${page}`);
  const scoped = indexed.items.filter((item) => page == null || item.pageId === page);
  const matches = query == null
    ? scoped.filter((item) => page != null || item.type === "page").map((item) => ({ ...item, score: 0 }))
    : scoped.map((item) => ({ ...item, score: fuzzyScore(query, item) })).filter((item) => item.score > 0);
  matches.sort((left, right) => right.score - left.score || left.order - right.order || left.id.localeCompare(right.id));
  const maximum = 100;
  return Object.freeze({
    ok: true,
    command: "inspect",
    input: workspace.path,
    programSha256: validated.programSha256,
    sourceBound: validated.sourceBound,
    sourceRevision: program.source?.revision ?? null,
    query: query ?? null,
    page: page ?? null,
    selectionMade: false,
    totalMatches: matches.length,
    truncated: matches.length > maximum,
    results: Object.freeze(matches.slice(0, maximum).map(({ order: _order, search: _search, ...item }) => Object.freeze(item))),
    diagnostics: validated.diagnostics,
  });
}

export async function checkPpj(
  { inputPath, fix = false },
  { cwd = process.cwd(), load = loadPpjWorkspace, validate = validatePpjWorkspace } = {},
) {
  const workspace = await load(inputPath, { cwd });
  const validated = await validate(workspace, { includeNodeMap: true });
  const formatted = prettyProgram(validated.programJson);
  const alreadyFormatted = Buffer.from(workspace.program).equals(formatted);
  if (fix && !alreadyFormatted) await replaceRegularFile(workspace.path, formatted);
  const program = JSON.parse(Buffer.from(validated.programJson).toString("utf8"));
  return Object.freeze({
    ok: true,
    command: "check",
    input: workspace.path,
    valid: true,
    fixed: Boolean(fix && !alreadyFormatted),
    canonical: true,
    formatted: Boolean(alreadyFormatted || fix),
    programSha256: validated.programSha256,
    sourceBound: validated.sourceBound,
    pages: Array.isArray(program.pages) ? program.pages.length : 0,
    expandedElementCount: validated.expandedElementCount,
    changedNodeIds: validated.changedNodeIds,
    diagnostics: validated.diagnostics,
  });
}

export async function buildPpj(
  { inputPath, outputPath },
  { cwd = process.cwd(), load = loadPpjWorkspace, compile = compilePpjWorkspace } = {},
) {
  const workspace = await load(inputPath, { cwd });
  const destination = path.resolve(cwd, outputPath);
  if (path.extname(destination).toLowerCase() !== ".pptx") throw new Error(`PPJ build output must be a .pptx file: ${destination}`);
  if (destination === workspace.path || destination === workspace.sourcePath) {
    throw new Error("PPJ build output must not overwrite the PPJ or its bound source PPTX.");
  }
  if (await pathExists(destination)) throw new Error(`PPTX output already exists: ${destination}`);
  const compiled = await compile(workspace, { includeNodeMap: true });
  if (!(compiled.file instanceof Uint8Array) || compiled.file.byteLength === 0) {
    throw new Error("OfficeKit native compiler returned no PPTX bytes.");
  }
  if (!/^[a-f0-9]{64}$/u.test(compiled.outputSha256) || sha256(compiled.file) !== compiled.outputSha256) {
    throw new Error("OfficeKit native compiler returned a PPTX with an invalid content hash.");
  }
  try {
    await writeExclusiveFile(destination, compiled.file, 0o644);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(`PPTX output already exists: ${destination}`);
    throw error;
  }
  return Object.freeze({
    ok: true,
    command: "build",
    input: workspace.path,
    output: destination,
    programSha256: compiled.programSha256,
    outputSha256: compiled.outputSha256,
    sourceBound: compiled.sourceBound,
    sourceSha256: compiled.sourceSha256 || null,
    expandedElementCount: compiled.expandedElementCount,
    changedParts: compiled.changedParts,
    changedNodeIds: compiled.changedNodeIds,
    diagnostics: compiled.diagnostics,
  });
}

function programPageCount(bytes) {
  const root = JSON.parse(Buffer.from(bytes).toString("utf8"));
  return Array.isArray(root.pages) ? root.pages.length : 0;
}

function indexProgram(program) {
  const items = [];
  const pages = new Map();
  let order = 0;
  for (const [pageIndex, page] of (program.pages ?? []).entries()) {
    pages.set(page.id, page);
    items.push(indexItem(page, { type: "page", pageId: page.id, pageIndex, zOrder: null, order: order++ }));
    for (const [zOrder, element] of (page.elements ?? []).entries()) {
      indexElement(element, { pageId: page.id, pageIndex, zOrder: [zOrder], order: () => order++ }, items);
    }
  }
  for (const component of program.components ?? []) {
    items.push(indexItem(component, { type: "component", pageId: null, pageIndex: null, zOrder: null, order: order++ }));
  }
  for (const asset of program.assets ?? []) {
    items.push(indexItem(asset, { type: "asset", pageId: null, pageIndex: null, zOrder: null, order: order++ }));
  }
  return { items, pages };
}

function indexElement(element, context, output) {
  output.push(indexItem(element, {
    type: element.type ?? "element",
    pageId: context.pageId,
    pageIndex: context.pageIndex,
    zOrder: context.zOrder,
    order: context.order(),
  }));
  if (!Array.isArray(element.children)) return;
  for (const [index, child] of element.children.entries()) {
    indexElement(child, { ...context, zOrder: [...context.zOrder, index] }, output);
  }
}

function indexItem(value, { type, pageId, pageIndex, zOrder, order }) {
  const summary = itemSummary(value);
  const capabilities = Array.isArray(value.nativeRef?.capabilities)
    ? value.nativeRef.capabilities.map((capability) => ({
      id: capability.id,
      operation: capability.operation,
      fields: capability.fields ?? [],
    }))
    : [];
  const frame = value.frame && typeof value.frame === "object" ? {
    x: value.frame.x,
    y: value.frame.y,
    width: value.frame.width,
    height: value.frame.height,
  } : null;
  const search = [value.id, value.name, value.role, value.claim, type, summary, ...capabilities.map(item => item.operation)]
    .filter((item) => typeof item === "string").join(" ");
  return {
    id: value.id,
    type,
    pageId,
    pageIndex,
    zOrder,
    name: value.name ?? null,
    role: value.role ?? null,
    summary,
    frame,
    sourceBound: Boolean(value.nativeRef),
    capabilities,
    score: 0,
    order,
    search,
  };
}

function itemSummary(value) {
  const strings = [];
  collectText(value.text, strings);
  if (Array.isArray(value.visibleText)) strings.push(...value.visibleText.filter((item) => typeof item === "string"));
  collectText(value.title, strings);
  collectText(value.data?.categories, strings);
  collectText(value.rows, strings);
  if (value.claim) strings.push(value.claim);
  if (value.description) strings.push(value.description);
  if (Array.isArray(value.categories)) strings.push(...value.categories.filter((item) => typeof item === "string"));
  const summary = strings.join(" ").replace(/\s+/gu, " ").trim();
  return summary.length > 240 ? `${summary.slice(0, 237)}...` : summary;
}

function collectText(value, output) {
  if (typeof value === "string") {
    output.push(value);
    return;
  }
  if (!value || typeof value !== "object") return;
  if (Array.isArray(value)) {
    for (const item of value) collectText(item, output);
    return;
  }
  if (typeof value.text === "string") output.push(value.text);
  if (Array.isArray(value.paragraphs)) collectText(value.paragraphs, output);
  if (Array.isArray(value.runs)) collectText(value.runs, output);
  if (Array.isArray(value.rows)) collectText(value.rows, output);
  if (Array.isArray(value.cells)) collectText(value.cells, output);
}

function fuzzyScore(query, item) {
  const needle = normalizeSearch(query);
  const haystack = normalizeSearch(item.search);
  if (!needle || !haystack) return 0;
  if (normalizeSearch(item.id) === needle) return 1_000;
  if (haystack.includes(needle)) return 700 - Math.min(200, haystack.indexOf(needle));
  const tokens = needle.split(" ").filter(Boolean);
  if (tokens.every((token) => haystack.includes(token))) return 500 + tokens.length;
  const matched = tokens.filter((token) => haystack.includes(token)).length;
  return matched === 0 ? 0 : 100 + matched;
}

function normalizeSearch(value) {
  return String(value ?? "").normalize("NFKC").toLocaleLowerCase("en-US").replace(/[^\p{L}\p{N}._:-]+/gu, " ").trim();
}

async function pathExists(target) {
  try {
    await lstat(target);
    return true;
  } catch (error) {
    if (error?.code === "ENOENT") return false;
    throw error;
  }
}

function formatResult(result) {
  if (result.command === "import") return [
    `OfficeKit imported ${result.input}`,
    `PPJ       ${result.output}`,
    `Source    ${result.source.path}`,
    `Revision  ${result.programSha256}`,
    `Pages     ${result.pages}`,
    `Assets    ${result.assets.length}`,
  ].join("\n");
  if (result.command === "check") return [
    `OfficeKit checked ${result.input}`,
    `Revision  ${result.programSha256}`,
    `Pages     ${result.pages}`,
    `Expanded  ${result.expandedElementCount}`,
    `Fixed     ${result.fixed ? "yes" : "no"}`,
  ].join("\n");
  if (result.command === "build") return [
    `OfficeKit built ${result.output}`,
    `Program   ${result.programSha256}`,
    `PPTX      ${result.outputSha256}`,
    `Changed   ${result.changedNodeIds.length} nodes / ${result.changedParts.length} parts`,
  ].join("\n");
  if (result.command === "inspect") {
    const lines = [
      `OfficeKit inspected ${result.input}`,
      `Revision  ${result.programSha256}`,
      `Matches   ${result.totalMatches}${result.truncated ? " (first 100 shown)" : ""}`,
    ];
    for (const item of result.results) {
      const location = item.pageId && item.type !== "page" ? ` · ${item.pageId}` : "";
      lines.push(`${item.id} · ${item.type}${location}${item.summary ? ` · ${item.summary}` : ""}`);
    }
    return lines.join("\n");
  }
  if (result.command === "render") return [
    `OfficeKit rendered ${result.pages.length} / ${result.pageCount} pages`,
    `Output    ${result.output}`,
    `Program   ${result.programSha256}`,
    `Renderer  ${result.renderer}`,
    `Review    ${result.visualReview}`,
  ].join("\n");
  if (result.command === "review") return [
    `OfficeKit reviewed ${result.input}`,
    `Program   ${result.programSha256}`,
    `Candidate ${result.candidateSha256}`,
    `Verdict   ${result.report.verdict}`,
    `Visual    ${result.visualReview}`,
    result.report.summary,
  ].join("\n");
  return JSON.stringify(result);
}
