#!/usr/bin/env node

/**
 * Read-only maintenance checker for the presentation primitive impact map.
 * It intentionally uses static text parsing rather than importing OfficeKit so
 * a Skill update cannot initialize a codec, provider, renderer, or bridge.
 */

import { execFileSync } from "node:child_process";
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import path from "node:path";
import process from "node:process";

const DEFAULT_MANIFEST = "skills/presentations/skills/presentations/references/primitive-impact.json";
const HELP_API_PATTERN = /\{\s*artifactKind:\s*"presentation"\s*,\s*kind:\s*"api"\s*,\s*name:\s*"([^"]+)"/g;
const PRIVATE_MARKERS = ["KIMI_API_KEY", "agent-gw", "kimi-slides", "presentation-artifact-tool", "/Users/"];

const args = process.argv.slice(2);
const command = args[0] && !args[0].startsWith("-") ? args.shift() : "check";
const options = parseOptions(args);
const repo = path.resolve(options.repo || process.cwd());
const manifestPath = path.join(repo, options.manifest || DEFAULT_MANIFEST);

const result = command === "impact"
  ? impact(repo, manifestPath, options)
  : check(repo, manifestPath);

print(result, options.json);
if (!result.ok) process.exitCode = 1;

function check(repoRoot, file) {
  const errors = [];
  const warnings = [];
  if (!existsSync(file)) return { ok: false, command: "check", errors: [`Missing impact manifest: ${path.relative(repoRoot, file)}`], warnings: [] };
  let manifest;
  try {
    manifest = JSON.parse(readFileSync(file, "utf8"));
  } catch (error) {
    return { ok: false, command: "check", errors: [`Invalid JSON: ${error.message}`], warnings: [] };
  }
  if (manifest.schema !== "office-kit/presentation-primitive-impact/v1") errors.push("Unexpected impact manifest schema");
  if (!Array.isArray(manifest.families) || manifest.families.length === 0) errors.push("Manifest must declare at least one family");

  const repoFiles = gitFiles(repoRoot);
  const helpSource = path.join(repoRoot, manifest.helpSource || "src/help/index.mjs");
  const helpNames = parseHelpNames(helpSource, errors);
  const owners = new Map(helpNames.map((name) => [name, []]));
  const familyIds = new Set();
  for (const family of manifest.families || []) {
    if (!family?.id || familyIds.has(family.id)) errors.push(`Family id is missing or duplicated: ${family?.id || "<empty>"}`);
    familyIds.add(family?.id);
    for (const pattern of family.helpPatterns || []) {
      if (!helpNames.some((name) => globMatch(name, pattern))) warnings.push(`${family.id}: help pattern matches no current API: ${pattern}`);
    }
    for (const name of helpNames) if ((family.helpPatterns || []).some((pattern) => globMatch(name, pattern))) owners.get(name).push(family.id);
    for (const field of ["runtimePaths", "protocolPaths", "consumerSkills", "references", "examples", "tests", "evidence"]) {
      for (const entry of family[field] || []) if (!repoFiles.some((candidate) => globMatch(candidate, entry))) errors.push(`${family.id}: path has no repository match: ${entry}`);
    }
  }
  for (const [name, familyList] of owners) if (familyList.length === 0) errors.push(`Presentation Help API is not mapped to a primitive family: ${name}`);
  for (const entry of manifest.globalSurfaces || []) if (!repoFiles.some((candidate) => globMatch(candidate, entry))) errors.push(`Global surface has no repository match: ${entry}`);

  const scanned = [
    path.join(repoRoot, "skills", "skill-update", "skills", "skill-update", "SKILL.md"),
    path.join(repoRoot, "skills", "presentations", "skills", "presentations", "references", "primitive-impact.json"),
  ];
  for (const filePath of scanned.flatMap((entry) => statSafe(entry)?.isDirectory() ? walk(entry) : [entry])) {
    if (!existsSync(filePath) || !statSafe(filePath)?.isFile()) continue;
    const source = readFileSync(filePath, "utf8");
    for (const marker of PRIVATE_MARKERS) if (source.includes(marker)) errors.push(`Clean-room/private marker in shipped guidance: ${path.relative(repoRoot, filePath)} contains ${marker}`);
  }
  return { ok: errors.length === 0, command: "check", manifest: path.relative(repoRoot, file), helpApiCount: helpNames.length, familyCount: manifest.families?.length || 0, errors, warnings };
}

function impact(repoRoot, file, options) {
  const manifest = readJson(file);
  const paths = options.paths?.length ? options.paths : changedFiles(repoRoot);
  const families = [];
  for (const family of manifest.families || []) {
    const matched = paths.filter((changed) => [
      ...(family.runtimePaths || []), ...(family.protocolPaths || []),
      ...(family.consumerSkills || []), ...(family.references || []),
      ...(family.examples || []), ...(family.tests || []), ...(family.evidence || []),
    ].some((pattern) => globMatch(changed, pattern)));
    if (matched.length) families.push({ id: family.id, summary: family.summary, changed: [...new Set(matched)], helpPatterns: family.helpPatterns || [], consumers: family.consumerSkills || [], references: family.references || [], examples: family.examples || [], tests: family.tests || [], evidence: family.evidence || [] });
  }
  const global = (manifest.globalSurfaces || []).filter((surface) => paths.some((changed) => globMatch(changed, surface)));
  return { ok: true, command: "impact", manifest: path.relative(repoRoot, file), changedPaths: paths, globalSurfaces: global, families, note: "Read-only impact report; review owners before editing." };
}

function parseOptions(argv) {
  const options = { paths: [] };
  for (let index = 0; index < argv.length; index += 1) {
    const value = argv[index];
    if (value === "--json") options.json = true;
    else if (value === "--repo" || value === "--manifest") options[value.slice(2)] = argv[++index];
    else if (value === "--paths") {
      while (argv[index + 1] && !argv[index + 1].startsWith("--")) options.paths.push(argv[++index]);
    } else throw new Error(`Unknown option: ${value}`);
  }
  return options;
}

function parseHelpNames(file, errors) {
  if (!existsSync(file)) { errors.push(`Missing Help source: ${path.relative(repo, file)}`); return []; }
  const text = readFileSync(file, "utf8");
  return [...text.matchAll(HELP_API_PATTERN)].map((match) => match[1]);
}

function readJson(file) { try { return JSON.parse(readFileSync(file, "utf8")); } catch (error) { throw new Error(`Cannot read ${file}: ${error.message}`); } }

function gitFiles(repoRoot) {
  try { return execFileSync("git", ["-C", repoRoot, "ls-files", "-co", "--exclude-standard"], { encoding: "utf8" }).split(/\r?\n/).filter(Boolean); }
  catch { return walk(repoRoot).map((file) => path.relative(repoRoot, file)); }
}

function changedFiles(repoRoot) {
  try {
    const tracked = execFileSync("git", ["-C", repoRoot, "diff", "--name-only", "HEAD", "--"], { encoding: "utf8" });
    const untracked = execFileSync("git", ["-C", repoRoot, "ls-files", "--others", "--exclude-standard"], { encoding: "utf8" });
    return [...new Set(`${tracked}\n${untracked}`.split(/\r?\n/).map((item) => item.trim()).filter(Boolean))];
  } catch { return []; }
}

function walk(root) {
  if (!existsSync(root)) return [];
  const output = [];
  for (const entry of readdirSync(root, { withFileTypes: true })) {
    if ([".git", "node_modules", "tmp", "output"].includes(entry.name)) continue;
    const target = path.join(root, entry.name);
    if (entry.isDirectory()) output.push(...walk(target));
    else if (entry.isFile()) output.push(target);
  }
  return output;
}

function statSafe(file) { try { return statSync(file); } catch { return null; } }

function globMatch(value, pattern) {
  const normalizedValue = String(value).replaceAll(path.sep, "/");
  const normalizedPattern = String(pattern).replaceAll(path.sep, "/");
  const escaped = normalizedPattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replaceAll("**", "§§").replaceAll("*", "[^/]*").replaceAll("§§", ".*");
  return new RegExp(`^${escaped}$`).test(normalizedValue);
}

function print(value, json) { process.stdout.write(json ? `${JSON.stringify(value, null, 2)}\n` : human(value)); }

function human(value) {
  if (value.command === "check") return [`Skill impact check: ${value.ok ? "ok" : "failed"}`, `Help APIs: ${value.helpApiCount ?? 0} · families: ${value.familyCount ?? 0}`, ...value.errors.map((item) => `ERROR ${item}`), ...value.warnings.map((item) => `WARN ${item}`)].join("\n") + "\n";
  const lines = [`Skill impact report: ${value.changedPaths.length} changed path(s)`];
  if (value.globalSurfaces.length) lines.push(`Global surfaces: ${value.globalSurfaces.join(", ")}`);
  for (const family of value.families) lines.push(`\n${family.id}\n  changed: ${family.changed.join(", ")}\n  consumers: ${family.consumers.join(", ")}\n  examples: ${family.examples.join(", ")}\n  tests: ${family.tests.join(", ")}\n  evidence: ${family.evidence.join(", ")}`);
  if (!value.families.length) lines.push("No mapped family; add the capability to the impact manifest before shipping.");
  return lines.join("\n") + "\n";
}
