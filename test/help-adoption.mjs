import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { HELP_CATALOG, queryHelpRecords } from "../src/help/index.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const presentation = HELP_CATALOG.filter((item) => item.artifactKind === "presentation");
assert.ok(presentation.length >= 100, "the public Presentation surface must be classified");
assert.ok(presentation.every((item) => ["golden", "advanced", "compatibility"].includes(item.adoptionTier)));
assert.ok(presentation.every((item) => ["golden", "advanced", "compatibility"].includes(item.adoptionTier)
  && item.useWhen?.length
  && item.avoidWhen?.length
  && item.requires?.length
  && item.review?.length
  && item.recipes?.length
  && item.examplePaths?.length));

const golden = presentation.filter((item) => item.adoptionTier === "golden");
assert.ok(golden.length >= 30, "golden adoption surface is unexpectedly small");
for (const item of golden) {
  for (const recipe of item.recipes) {
    const recipePath = recipe.split("#", 1)[0];
    await fs.access(path.join(repoRoot, recipePath));
  }
  for (const examplePath of item.examplePaths) await fs.access(path.join(repoRoot, examplePath));
  assert.ok(item.examples?.length || item.examplePaths?.length, `${item.name} needs an example pointer`);
}

const composeResults = queryHelpRecords("presentation", "compose", { adoptionTier: "golden", search: "reader outcome|authoring plan" });
assert.ok(composeResults.some((item) => item.name === "slide.compose"));
assert.ok(composeResults.every((item) => item.adoptionTier === "golden"));

const reuseResults = queryHelpRecords("presentation", "reuse", { adoptionTier: "golden", search: "source-bound|inspect-backed" });
assert.ok(reuseResults.some((item) => item.name === "presentation.reuseSourceSlide"));
assert.ok(reuseResults.some((item) => item.name === "presentation.reuseSourceComponent"));

const advancedResults = queryHelpRecords("presentation", "comments", { adoptionTier: "advanced" });
assert.ok(advancedResults.some((item) => item.name === "slide.comments.addThread"));
assert.ok(advancedResults.every((item) => item.adoptionTier === "advanced"));

console.log(`Help adoption ok: ${presentation.length} Presentation records, ${golden.length} golden`);
