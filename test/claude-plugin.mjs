import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const marketplacePath = path.join(repoRoot, ".claude-plugin", "marketplace.json");
const marketplace = JSON.parse(await fs.readFile(marketplacePath, "utf8"));

assert.equal(marketplace.name, "officekit");
assert.match(marketplace.description, /Office and PDF/i);
assert.equal(marketplace.owner.name, "w31r4");
assert.equal(marketplace.metadata.version, "1.0.0");
assert.ok(Array.isArray(marketplace.plugins));

const expected = new Map([
  ["office-kit", { version: "1.0.0", license: "AGPL-3.0-or-later" }],
  ["documents", { version: "0.2.0", license: "AGPL-3.0-or-later" }],
  ["spreadsheets", { version: "0.2.0", license: "AGPL-3.0-or-later" }],
  ["presentations", { version: "0.2.0", license: "AGPL-3.0-or-later" }],
  ["pdf", { version: "0.2.0", license: "AGPL-3.0-or-later" }],
  ["template-creator", { version: "0.2.0", license: "AGPL-3.0-or-later" }],
  ["presentation-template-creator", { version: "1.1.0", license: "AGPL-3.0-or-later" }],
  ["presentation-template-library", { version: "1.1.0", license: "AGPL-3.0-or-later" }],
  ["default-template-library", { version: "0.2.0", license: "MIT" }],
]);
assert.deepEqual(
  marketplace.plugins.map((plugin) => plugin.name),
  [...expected.keys()],
  "Claude marketplace plugin order is part of the discovery surface",
);

for (const plugin of marketplace.plugins) {
  const metadata = expected.get(plugin.name);
  assert.ok(metadata, `unexpected Claude plugin ${plugin.name}`);
  assert.equal(plugin.version, metadata.version);
  assert.equal(plugin.strict, false, `${plugin.name} must use the marketplace definition`);
  assert.equal(plugin.skills, "./skills/");
  assert.equal(plugin.license, metadata.license);
  assert.match(plugin.source, /^\.\/skills\/[a-z0-9-]+$/);
  assert.doesNotMatch(plugin.source, /\.\./);
  const pluginRoot = path.resolve(repoRoot, plugin.source);
  const relativePluginRoot = path.relative(repoRoot, pluginRoot);
  assert.ok(relativePluginRoot && !relativePluginRoot.startsWith(".."));
  const pluginStat = await fs.stat(pluginRoot);
  assert.ok(pluginStat.isDirectory(), `${plugin.name} source must be a directory`);

  const skillsRoot = path.resolve(pluginRoot, plugin.skills);
  assert.equal(path.relative(pluginRoot, skillsRoot), "skills");
  const skillEntries = (await fs.readdir(skillsRoot, { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();
  assert.ok(skillEntries.length > 0, `${plugin.name} must expose at least one Skill`);
  for (const skillName of skillEntries) {
    const skillPath = path.join(skillsRoot, skillName, "SKILL.md");
    const skillText = await fs.readFile(skillPath, "utf8");
    assert.match(skillText, /^---\n[\s\S]*?\n---/);
  }
}

console.log("Claude plugin marketplace smoke ok");
