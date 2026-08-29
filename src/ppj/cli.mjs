import { createHash, randomUUID } from "node:crypto";
import {
  chmod,
  link,
  lstat,
  mkdir,
  readFile,
  realpath,
  rm,
  writeFile,
} from "node:fs/promises";
import path from "node:path";
import process from "node:process";

import { projectPptxToPpj } from "./native.mjs";

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
  if (subcommand !== "import") {
    throw new Error(`Unknown or unavailable PPJ command "${subcommand}". Run "officekit ppj --help".`);
  }
  const request = parseImportArguments(rest);
  if (request.help) {
    output.write(`${PPJ_USAGE}\n`);
    return;
  }
  const result = await importPptxAsPpj(request, { cwd });
  output.write(request.json ? `${JSON.stringify(result)}\n` : `${formatImportResult(result)}\n`);
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

function requiredValue(args, index, option) {
  const value = args[index];
  if (!value || value.startsWith("-")) throw new Error(`${option} requires a value.`);
  return value;
}

export async function importPptxAsPpj(
  { inputPath, outputPath },
  { cwd = process.cwd(), project = projectPptxToPpj } = {},
) {
  const input = await regularFile(path.resolve(cwd, inputPath), "PPTX input");
  const destination = path.resolve(cwd, outputPath);
  if (path.extname(input).toLowerCase() !== ".pptx") throw new Error(`PPJ import input must be a .pptx file: ${input}`);
  if (path.extname(destination).toLowerCase() !== ".ppj") throw new Error(`PPJ import output must be a .ppj file: ${destination}`);
  if (input === destination) throw new Error("PPJ import cannot overwrite its PPTX input.");
  if (await statOrNull(destination)) throw new Error(`PPJ output already exists: ${destination}`);

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
  await writeExclusive(destination, projected.programJson, 0o644);
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

async function regularFile(target, label) {
  const resolved = await realpath(target).catch((error) => {
    if (error?.code === "ENOENT") throw new Error(`${label} does not exist: ${target}`);
    throw error;
  });
  const stat = await lstat(resolved);
  if (!stat.isFile() || stat.isSymbolicLink()) throw new Error(`${label} must be a regular file: ${resolved}`);
  return resolved;
}

async function writeImmutableContent(target, data, expectedSha256) {
  const existing = await statOrNull(target);
  if (existing) {
    if (!existing.isFile() || existing.isSymbolicLink() || sha256(await readFile(target)) !== expectedSha256) {
      throw new Error(`Content-addressed PPJ asset conflicts with existing path: ${target}`);
    }
    await chmod(target, 0o444);
    return;
  }
  await mkdir(path.dirname(target), { recursive: true });
  await writeExclusive(target, data, 0o444);
}

async function writeExclusive(target, data, mode) {
  const temporary = path.join(path.dirname(target), `.${path.basename(target)}.${process.pid}-${randomUUID()}.tmp`);
  try {
    await writeFile(temporary, data, { flag: "wx", mode });
    await chmod(temporary, mode);
    await link(temporary, target);
  } finally {
    await rm(temporary, { force: true });
  }
}

async function statOrNull(target) {
  try {
    return await lstat(target);
  } catch (error) {
    if (error?.code === "ENOENT") return null;
    throw error;
  }
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function programPageCount(bytes) {
  const root = JSON.parse(Buffer.from(bytes).toString("utf8"));
  return Array.isArray(root.pages) ? root.pages.length : 0;
}

function formatImportResult(result) {
  return [
    `OfficeKit imported ${result.input}`,
    `PPJ       ${result.output}`,
    `Source    ${result.source.path}`,
    `Revision  ${result.programSha256}`,
    `Pages     ${result.pages}`,
    `Assets    ${result.assets.length}`,
  ].join("\n");
}
