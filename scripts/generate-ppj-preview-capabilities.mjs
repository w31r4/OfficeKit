import { readFile, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { derivePreviewCapabilities } from "../src/ppj/preview-capabilities.mjs";

const args = process.argv.slice(2);
if (args.some((arg) => arg !== "--check") || args.length > 1) throw new Error("Usage: node scripts/generate-ppj-preview-capabilities.mjs [--check]");
const root = new URL("../", import.meta.url);
const registry = JSON.parse(await readFile(new URL("src/ppj/capability-registry.json", root), "utf8"));
const schema = JSON.parse(await readFile(new URL("src/ppj/ppj-v1.schema.json", root), "utf8"));
const target = new URL("src/ppj/svg-preview-capabilities.json", root);
const content = `${JSON.stringify(derivePreviewCapabilities(registry, schema), null, 2)}\n`;
if (args.includes("--check")) {
  if (await readFile(target, "utf8") !== content) throw new Error("Preview capability summary is stale; run node scripts/generate-ppj-preview-capabilities.mjs");
  console.log("ppj preview capability summary is current");
} else {
  await writeFile(target, content);
  console.log(`wrote ${fileURLToPath(target)}`);
}
