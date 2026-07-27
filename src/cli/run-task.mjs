import { lstat, readFile, realpath } from "node:fs/promises";
import Module, { registerHooks } from "node:module";
import path from "node:path";
import process from "node:process";
import { fileURLToPath, pathToFileURL } from "node:url";

const PACKAGE_ROOT = fileURLToPath(new URL("../..", import.meta.url));
const TASK_EXTENSIONS = new Set([".js", ".mjs"]);

export const RUN_TASK_USAGE = [
  "Usage: officekit run <task.mjs> [-- <task arguments>]",
  "",
  "Run a local JavaScript task with this OfficeKit installation.",
  "Imports of office-kit and its public subpaths resolve to this CLI version.",
].join("\n");

export async function runTaskCommand(args, { output = process.stdout } = {}) {
  if (args.length === 0 || args[0] === "--help" || args[0] === "-h") {
    output.write(`${RUN_TASK_USAGE}\n`);
    return;
  }
  const [requestedScript, ...remainder] = args;
  if (
    requestedScript === "-" ||
    requestedScript.includes("\0") ||
    /^[a-z][a-z0-9+.-]*:/iu.test(requestedScript)
  ) {
    throw new Error("officekit run accepts one local .mjs or .js file path, not stdin or a URL.");
  }
  const taskArguments = remainder[0] === "--" ? remainder.slice(1) : remainder;
  const absoluteScript = path.resolve(requestedScript);
  const scriptStat = await lstat(absoluteScript).catch((error) => {
    if (error?.code === "ENOENT") return null;
    throw error;
  });
  if (scriptStat == null || !scriptStat.isFile()) {
    throw new Error(`OfficeKit task is not a regular file: ${absoluteScript}`);
  }
  if (!TASK_EXTENSIONS.has(path.extname(absoluteScript).toLowerCase())) {
    throw new Error("officekit run accepts only .mjs or .js task files.");
  }

  const [canonicalScript, packageMetadata] = await Promise.all([
    realpath(absoluteScript),
    readPackageMetadata(),
  ]);
  const exportTargets = exportedOfficeKitTargets(packageMetadata);
  const commonJsTargets = new Map(
    [...exportTargets].map(([specifier, target]) => [
      specifier,
      fileURLToPath(target),
    ]),
  );
  const unpublishedSubpathError = (specifier) =>
    new Error(`OfficeKit task requested unpublished package subpath "${specifier}".`);
  const isOfficeKitSpecifier = (specifier) =>
    specifier === packageMetadata.name ||
    specifier.startsWith(`${packageMetadata.name}/`);
  const hooks = registerHooks({
    resolve(specifier, context, nextResolve) {
      const target = exportTargets.get(specifier);
      if (target != null) return { shortCircuit: true, url: target };
      if (isOfficeKitSpecifier(specifier)) throw unpublishedSubpathError(specifier);
      return nextResolve(specifier, context);
    },
  });
  const originalCommonJsResolve = Module._resolveFilename;
  const commonJsResolve = function resolveOfficeKitFromCli(
    specifier,
    parent,
    isMain,
    options,
  ) {
    const target = commonJsTargets.get(specifier);
    if (target != null) return target;
    if (isOfficeKitSpecifier(specifier)) throw unpublishedSubpathError(specifier);
    return Reflect.apply(originalCommonJsResolve, this, [
      specifier,
      parent,
      isMain,
      options,
    ]);
  };
  Module._resolveFilename = commonJsResolve;

  const originalArgv = process.argv;
  process.argv = [process.execPath, canonicalScript, ...taskArguments];
  try {
    await import(pathToFileURL(canonicalScript).href);
  } catch (error) {
    if (error != null && (typeof error === "object" || typeof error === "function")) {
      Object.defineProperty(error, "officeKitShowStack", {
        configurable: true,
        value: true,
      });
    }
    throw error;
  } finally {
    process.argv = originalArgv;
    Module._resolveFilename = originalCommonJsResolve;
    hooks.deregister();
  }
}

async function readPackageMetadata() {
  const metadata = JSON.parse(
    await readFile(path.join(PACKAGE_ROOT, "package.json"), "utf8"),
  );
  if (
    typeof metadata.name !== "string" ||
    metadata.exports == null ||
    typeof metadata.exports !== "object" ||
    Array.isArray(metadata.exports)
  ) {
    throw new Error("OfficeKit package metadata does not expose a valid exports map.");
  }
  return metadata;
}

function exportedOfficeKitTargets(metadata) {
  const result = new Map();
  for (const [subpath, target] of Object.entries(metadata.exports)) {
    if (
      (subpath !== "." && !subpath.startsWith("./")) ||
      typeof target !== "string" ||
      !target.startsWith("./")
    ) {
      continue;
    }
    const specifier = subpath === "."
      ? metadata.name
      : `${metadata.name}/${subpath.slice(2)}`;
    const absoluteTarget = path.resolve(PACKAGE_ROOT, target);
    const relative = path.relative(PACKAGE_ROOT, absoluteTarget);
    if (
      relative === "" ||
      relative === ".." ||
      relative.startsWith(`..${path.sep}`) ||
      path.isAbsolute(relative)
    ) {
      throw new Error(`OfficeKit export "${subpath}" escapes the package root.`);
    }
    result.set(specifier, pathToFileURL(absoluteTarget).href);
  }
  return result;
}
