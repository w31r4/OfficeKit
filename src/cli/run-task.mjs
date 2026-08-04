import { lstat, realpath } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { pathToFileURL } from "node:url";
import {
  createOfficeKitResolver,
  installOfficeKitModuleHooks,
  readOfficeKitPackageMetadata,
} from "./officekit-resolver.mjs";

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
  const urlScheme = /^[a-z][a-z0-9+.-]*:/iu.test(requestedScript);
  const windowsDrivePath =
    process.platform === "win32" && /^[a-z]:/iu.test(requestedScript);
  if (
    requestedScript === "-" ||
    requestedScript.includes("\0") ||
    (urlScheme && !windowsDrivePath)
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
    readOfficeKitPackageMetadata(),
  ]);
  const resolver = createOfficeKitResolver(packageMetadata);
  const deregisterHooks = installOfficeKitModuleHooks(resolver);

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
    deregisterHooks();
  }
}
