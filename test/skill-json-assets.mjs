import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import {
  COMPACT_SKILL_JSON_PATHS,
  REPO_ROOT,
  compactJsonText,
  compactSkillJsons,
} from "../scripts/compact-skill-jsons.mjs";

assert.deepEqual(COMPACT_SKILL_JSON_PATHS, [
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library/artifact-tool-compose/content-tokens.json",
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library/artifact-tool-compose/template-registry.json",
]);
assert.equal(compactJsonText('{\n  "a": 1,\n  "b": [true, null]\n}\n'), '{"a":1,"b":[true,null]}\n');
assert.equal(compactJsonText('{"emoji":"栗子"}\n'), '{"emoji":"栗子"}\n');
assert.throws(() => compactJsonText("{", "broken fixture"), /broken fixture is not valid JSON/);
assert.throws(() => compactJsonText("x".repeat(2 * 1024 * 1024 + 1)), /source budget/);

let totalBytes = 0;
for (const relativePath of COMPACT_SKILL_JSON_PATHS) {
  const filename = path.join(REPO_ROOT, relativePath);
  const stat = await fs.lstat(filename);
  assert.equal(stat.isFile(), true, `${relativePath} must remain a regular file`);
  const source = await fs.readFile(filename, "utf8");
  assert.equal(source, compactJsonText(source, relativePath), `${relativePath} must remain deterministically compact`);
  totalBytes += Buffer.byteLength(source);
}
assert.ok(totalBytes <= 335_000, `derived Skill JSON assets exceed the 335,000-byte compact budget (${totalBytes})`);
assert.deepEqual(await compactSkillJsons(), { files: 2, changed: 0, totalBytes, savings: 0 });

console.log(`derived Skill JSON asset integrity ok: 2 files, ${totalBytes} bytes`);
