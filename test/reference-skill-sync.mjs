import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";

import {
  checkReferenceSkillSync,
  createReferenceSkillSnapshot,
  REFERENCE_SKILL_BUNDLES,
  REPLACED_REFERENCE_SKILL_PATHS,
  RETIRED_REFERENCE_SKILL_PATH_PREFIXES,
} from "../scripts/reference-skill-sync.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const checked = await checkReferenceSkillSync();
const rebuilt = await createReferenceSkillSnapshot();
const recorded = JSON.parse(await fs.readFile(path.join(repoRoot, "skills", "reference-sync.json"), "utf8"));

assert.deepEqual(checked, rebuilt);
assert.deepEqual(rebuilt, recorded);
assert.equal(rebuilt.schemaVersion, 1);
assert.equal(rebuilt.source.commit, "73c99c67ca7bbaa82cec0b158c647db583dcd970");
assert.equal(rebuilt.totalFiles, 333);
assert.deepEqual(Object.keys(rebuilt.bundles), REFERENCE_SKILL_BUNDLES);
assert.equal(Object.values(rebuilt.bundles).reduce((sum, bundle) => sum + bundle.files, 0), rebuilt.totalFiles);
assert.equal(Object.values(rebuilt.bundles).reduce((sum, bundle) => sum + bundle.bytes, 0), rebuilt.totalBytes);
assert.deepEqual([...REPLACED_REFERENCE_SKILL_PATHS].sort(), [
  "spreadsheets/.app.json",
  "spreadsheets/skills/excel-live-control/officejs.md",
]);
assert.deepEqual(RETIRED_REFERENCE_SKILL_PATH_PREFIXES, [
  "presentations/skills/presentations/assets/builtin_templates/grid-layout-library/",
  "presentations/skills/presentations/builtin_templates_support/",
  "default-template-library/skills/artifact-template-business-review/",
  "default-template-library/skills/artifact-template-market-trends-report/",
  "default-template-library/skills/artifact-template-operating-review/",
  "default-template-library/skills/artifact-template-project-kickoff/",
  "default-template-library/skills/artifact-template-simple-dark-mode/",
  "default-template-library/skills/artifact-template-simple-light-mode/",
  "default-template-library/skills/artifact-template-team-alignment/",
]);

const referenceChecklist = await fs.readFile(path.join(
  repoRoot,
  "reference",
  "office-artifact-tool",
  "skills",
  "documents",
  "skills",
  "documents",
  "examples",
  "end_to_end_smoke_test.md",
));
const projectChecklist = await fs.readFile(path.join(
  repoRoot,
  "skills",
  "documents",
  "skills",
  "documents",
  "examples",
  "end_to_end_smoke_test.md",
));
assert.deepEqual(projectChecklist, referenceChecklist, "the newly synchronized reference checklist must remain byte-identical");

console.log(`reference Skill source sync ok: ${rebuilt.totalFiles} files at ${rebuilt.source.commit}`);
