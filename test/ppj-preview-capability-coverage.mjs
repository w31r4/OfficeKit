import assert from "node:assert/strict";
import registry from "../src/ppj/capability-registry.json" with { type: "json" };
import preview from "../src/ppj/svg-preview-capabilities.json" with { type: "json" };
import { readdir, readFile } from "node:fs/promises";
import path from "node:path";

const declared = new Set([...preview.supported, ...preview.partial, ...preview.opaque]);
for (const boundary of registry.authoredCompilerBoundaries) {
  const text = `${boundary.feature} ${boundary.ppjPath}`.toLowerCase();
  for (const chart of ["heatmap", "treemap", "sunburst", "sankey", "waterfall", "candlestick", "streamgraph", "pictographic", "combo", "scatter", "bubble", "radar", "pie", "doughnut", "bar", "column", "line", "area"]) {
    if (text.includes(chart)) assert.ok(declared.has(`chart:${chart}`) || declared.has(chart), `renderer capability missing declaration for chart:${chart}`);
  }
}
console.log(`ppj preview capability coverage ok (${declared.size} declared renderer capabilities)`);

const fixtureRoot = path.resolve("test/fixtures");
async function walk(dir) {
  const out = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const file = path.join(dir, entry.name);
    if (entry.isDirectory()) out.push(...await walk(file));
    else if (entry.name.endsWith(".ppj")) out.push(file);
  }
  return out;
}
for (const file of await walk(fixtureRoot)) {
  const program = JSON.parse(await readFile(file, "utf8"));
  for (const page of program.pages || []) for (const element of page.elements || []) {
    assert.ok(declared.has(element.type) || element.type === "chart", `renderer capability missing declaration for ${element.type} (${file})`);
  }
}
