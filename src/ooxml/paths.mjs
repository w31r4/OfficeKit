import path from "node:path";

export function ooxmlSafePartPath(partPath, family = "OOXML") {
  const raw = String(partPath || "").replaceAll("\\", "/").trim();
  if (!raw || raw.startsWith("/") || raw.includes("\0")) throw new Error(`Unsafe ${family} part path: ${partPath}`);
  const normalized = path.posix.normalize(raw).replace(/^\.\//, "");
  if (!normalized || normalized === "." || normalized.startsWith("../") || normalized.includes("/../") || normalized === "..") {
    throw new Error(`Unsafe ${family} part path: ${partPath}`);
  }
  if (normalized.length > 1024) throw new Error(`Unsafe ${family} part path: path exceeds 1024 characters`);
  return normalized;
}

export function ooxmlResolveRelationshipTarget(source, rawTarget) {
  const target = String(rawTarget || "").split("#")[0];
  if (target.startsWith("/")) return target.slice(1);
  const sourceDir = source ? path.posix.dirname(source) : "";
  return path.posix.normalize(path.posix.join(sourceDir === "." ? "" : sourceDir, target));
}
