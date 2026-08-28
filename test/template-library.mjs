import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";

const legacyRoots = [
  path.resolve(import.meta.dirname, "../skills/presentations/skills/presentations/assets/builtin_templates"),
  path.resolve(import.meta.dirname, "../skills/presentations/skills/presentations/builtin_templates_support"),
];

for (const root of legacyRoots) {
  await assert.rejects(
    fs.access(root),
    (error) => error?.code === "ENOENT",
    `legacy built-in presentation assets must be retired: ${root}`,
  );
}

console.log("legacy presentation template assets retired");
