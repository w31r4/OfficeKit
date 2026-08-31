import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { PUBLIC_HELP_CATALOG, queryHelpRecords } from "../src/help/index.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const presentation = PUBLIC_HELP_CATALOG.filter((item) => item.artifactKind === "presentation");
const expected = [
  "officekit ppj resume",
  "officekit ppj import",
  "officekit ppj inspect",
  "officekit ppj check",
  "officekit ppj build",
  "officekit ppj render",
  "officekit ppj review",
];

assert.deepEqual(presentation.map((item) => item.name), expected);
for (const item of presentation) {
  assert.equal(item.adoptionTier, "golden");
  assert.ok(item.useWhen?.length && item.avoidWhen?.length && item.requires?.length && item.review?.length);
  assert.deepEqual(item.examplePaths, ["skills/presentations/skills/presentations/references/ppj.md"]);
  await fs.access(path.join(repoRoot, item.examplePaths[0]));
}

const sourceResults = queryHelpRecords("presentation", "source-bound");
assert.ok(sourceResults.some((item) => item.name === "officekit ppj build"));
assert.ok(sourceResults.every((item) => item.name.startsWith("officekit ppj ")));
assert.equal(queryHelpRecords("presentation", "slide.compose").length, 0);
assert.equal(queryHelpRecords("presentation", "PresentationFile.importPptx").length, 0);

console.log("Help adoption ok: PPJ is the only public Presentation route");
