import { createRequire, registerHooks } from "node:module";
import Module from "node:module";
import { realpathSync } from "node:fs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

export const PACKAGE_ROOT = fileURLToPath(new URL("../..", import.meta.url));

export async function readOfficeKitPackageMetadata() {
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

export function createOfficeKitResolver(metadata) {
  const exportTargets = exportedOfficeKitTargets(metadata);
  const commonJsTargets = new Map(
    [...exportTargets].map(([specifier, target]) => [specifier, fileURLToPath(target)]),
  );
  const isOfficeKitSpecifier = (specifier) =>
    specifier === metadata.name || specifier.startsWith(`${metadata.name}/`);
  const unpublishedSubpathError = (specifier) => {
    const error = new Error(`OfficeKit task requested unpublished package subpath "${specifier}".`);
    error.code = "unpublished-subpath";
    return error;
  };

  return Object.freeze({
    metadata,
    exportTargets,
    commonJsTargets,
    isOfficeKitSpecifier,
    unpublishedSubpathError,
    resolvePublished(specifier) {
      const target = exportTargets.get(specifier);
      if (target != null) return target;
      if (isOfficeKitSpecifier(specifier)) throw unpublishedSubpathError(specifier);
      return null;
    },
    resolvePublishedPath(specifier) {
      const target = commonJsTargets.get(specifier);
      if (target != null) return target;
      if (isOfficeKitSpecifier(specifier)) throw unpublishedSubpathError(specifier);
      return null;
    },
  });
}

export function installOfficeKitModuleHooks(resolver) {
  const hooks = registerHooks({
    resolve(specifier, context, nextResolve) {
      const target = resolver.resolvePublished(specifier);
      if (target != null) return { shortCircuit: true, url: target };
      return nextResolve(specifier, context);
    },
  });
  const originalCommonJsResolve = Module._resolveFilename;
  Module._resolveFilename = function resolveOfficeKitFromCli(
    specifier,
    parent,
    isMain,
    options,
  ) {
    const target = resolver.resolvePublishedPath(specifier);
    if (target != null) return target;
    return Reflect.apply(originalCommonJsResolve, this, [
      specifier,
      parent,
      isMain,
      options,
    ]);
  };
  return () => {
    Module._resolveFilename = originalCommonJsResolve;
    hooks.deregister();
  };
}

export function resolveWorkspaceSpecifier(
  resolver,
  specifier,
  { workspaceRoot, allowRelative = true } = {},
) {
  if (typeof specifier !== "string" || specifier.length === 0) {
    throw new TypeError("OfficeKit import specifier must be a non-empty string.");
  }
  if (resolver.exportTargets.has(specifier)) return resolver.exportTargets.get(specifier);
  if (resolver.isOfficeKitSpecifier(specifier)) {
    throw resolver.unpublishedSubpathError(specifier);
  }
  if (specifier.startsWith("node:")) return specifier;
  if (/^[a-z][a-z0-9+.-]*:/iu.test(specifier)) {
    const error = new Error("OfficeKit REPL imports accept local modules, not URLs.");
    error.code = "remote-import";
    throw error;
  }
  if (workspaceRoot == null) {
    throw new Error("OfficeKit REPL requires a workspace root for local imports.");
  }
  const root = path.resolve(workspaceRoot);
  if (
    specifier.startsWith(".") ||
    specifier.startsWith("/") ||
    path.win32.isAbsolute(specifier)
  ) {
    if (!allowRelative || path.isAbsolute(specifier) || path.win32.isAbsolute(specifier)) {
      const error = new Error("OfficeKit REPL imports reject absolute module paths.");
      error.code = "unsafe-import";
      throw error;
    }
    const candidate = path.resolve(root, specifier);
    const relative = path.relative(root, candidate);
    if (
      relative === ".." ||
      relative.startsWith(`..${path.sep}`) ||
      path.isAbsolute(relative)
    ) {
      const error = new Error("OfficeKit REPL import escapes workspaceRoot.");
      error.code = "unsafe-import";
      throw error;
    }
    return pathToFileURL(assertWorkspacePath(candidate, root)).href;
  }
  const require = createRequire(path.join(root, "package.json"));
  const resolved = require.resolve(specifier);
  return pathToFileURL(assertWorkspacePath(resolved, root)).href;
}

function assertWorkspacePath(candidate, root) {
  const canonical = realpathSync(candidate);
  const relative = path.relative(root, canonical);
  if (
    relative === ".." ||
    relative.startsWith(`..${path.sep}`) ||
    path.isAbsolute(relative)
  ) {
    const error = new Error("OfficeKit REPL import resolves outside workspaceRoot.");
    error.code = "unsafe-import";
    throw error;
  }
  return canonical;
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
