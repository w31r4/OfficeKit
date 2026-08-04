import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { queryTemplates } from "../src/templates/search.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const pluginRoot = path.join(repoRoot, "skills", "office-kit");
const skillRoot = path.join(pluginRoot, "skills", "office-kit");
const templateRoot = path.join(repoRoot, "skills", "default-template-library", "skills");
const officeKitCli = path.join(repoRoot, "bin", "officekit.mjs");

const [plugin, skillText, agentText, routingText, templateSelectionText] = await Promise.all([
  readJson(path.join(pluginRoot, ".codex-plugin", "plugin.json")),
  fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8"),
  fs.readFile(path.join(skillRoot, "agents", "openai.yaml"), "utf8"),
  fs.readFile(path.join(skillRoot, "references", "routing.md"), "utf8"),
  fs.readFile(path.join(skillRoot, "references", "template-selection.md"), "utf8"),
]);

assert.equal(plugin.name, "office-kit");
assert.equal(plugin.version, "0.6.0");
assert.equal(plugin.license, "AGPL-3.0-or-later");
assert.equal(plugin.skills, "./skills/");
assert.match(plugin.description, /cross-format Office and PDF/i);
assert.match(skillText, /^---\nname: office-kit\ndescription: .+broad, ambiguous, cross-format/m);
assert.doesNotMatch(skillText, /\[TODO:/);
assert.ok(skillText.split(/\r?\n/).length < 180, "OfficeKit SKILL.md must stay compact");
assert.match(skillText, /exactly one owning Skill/i);
assert.match(skillText, /selected`, `ask`, or `none`/);
assert.match(skillText, /Do not preload every Office Skill/);
assert.match(skillText, /Do not send the user away to repeat the request/);
assert.match(skillText, /PDF-only task/);
assert.match(skillText, /officekit template search \.\.\. --json/);
assert.match(skillText, /English search terms/);
assert.doesNotMatch(skillText, /query-templates\.mjs/);
assert.match(agentText, /display_name: "OfficeKit"/);
assert.match(agentText, /default_prompt: "Use \$office-kit /);
assert.match(routingText, /\.\.\/documents\/SKILL\.md/);
assert.match(routingText, /Excel Live Control/);
assert.match(routingText, /Spreadsheets -> Presentations -> PDF/);
assert.match(templateSelectionText, /copy-only/);
assert.match(templateSelectionText, /不用模板，由领域 Skill 设计/);
assert.match(templateSelectionText, /does not select a\s+template/i);
assert.match(templateSelectionText, /untrusted descriptive data/i);
assert.match(templateSelectionText, /attribution, not permission to access the network/i);
assert.match(templateSelectionText, /Classify the current task, not the user/i);
assert.match(templateSelectionText, /Clear \| Not specified \| Query the catalog/i);
assert.match(templateSelectionText, /Unclear \| Specified or uploaded/i);
assert.match(templateSelectionText, /task-scoped user reference/i);
assert.match(templateSelectionText, /must not be copied into a template directory/i);
assert.match(templateSelectionText, /BM25F remains deterministic and local/i);
assert.match(templateSelectionText, /does not call a model, build a vector/i);
assert.match(templateSelectionText, /one English canonical form/i);
assert.match(templateSelectionText, /templates bundled with OfficeKit/i);

const expectedCounts = new Map([
  ["document", 7],
  ["presentation", 7],
  ["spreadsheet", 6],
]);
for (const [kind, expectedCount] of expectedCounts) {
  const result = await queryTemplates({
    kind,
    roots: [templateRoot],
    maxCandidates: 20,
  });
  assert.equal(result.candidates.length, expectedCount, `${kind} template count`);
  assert.deepEqual(result.invalid, [], `${kind} template metadata`);
  assert.deepEqual(result.rejected, []);
  assert.equal(result.schemaVersion, 2);
  assert.equal(result.ranking.algorithm, "bm25f");
  assert.deepEqual(result.ranking.queryTerms, []);
  assert.equal(result.selectionMade, false);
  for (const candidate of result.candidates) {
    assert.equal(candidate.kind, kind);
    assert.ok(candidate.useWhen.length > 0);
    assert.ok(["neutral", "opinionated"].includes(candidate.visualCommitment));
    assert.ok(["copy-only", "bounded-edit", "composable"].includes(candidate.editProfile.level));
    assert.equal(candidate.provenance.license, "MIT");
    assert.equal(path.basename(candidate.skillPath), "SKILL.md");
    await Promise.all([
      fs.access(candidate.skillPath),
      fs.access(candidate.referencePath),
      fs.access(candidate.previewPath),
    ]);
  }
}

const ranked = await queryTemplates({
  kind: "presentation",
  roots: [templateRoot],
  tags: ["executive", "quarterly"],
  maxCandidates: 3,
});
assert.equal(ranked.candidates[0].id, "artifact-template-business-review");
assert.deepEqual(ranked.candidates[0].matchedTags, ["executive", "quarterly"]);
assert.equal(ranked.candidates.length, 3);
assert.equal(ranked.ranking.algorithm, "bm25f");
assert.ok(ranked.candidates[0].match.bm25 > ranked.candidates[1].match.bm25);

const structuredRanked = await queryTemplates({
  kind: "presentation",
  roots: [templateRoot],
  intent: {
    purposes: ["quarterly business review"],
    audiences: ["executive"],
    contentShapes: ["KPIs", "decisions"],
    visualTraits: {
      tone: ["formal"],
      density: "medium",
      colorMode: "neutral",
      structure: ["sectioned"],
    },
    requiredOperations: ["recognized-placeholder-title-text-replace"],
    brandSensitive: false,
  },
  maxCandidates: 5,
});
assert.equal(
  structuredRanked.candidates[0].id,
  "artifact-template-business-review",
);
assert.equal(structuredRanked.candidates[0].match.score, 100);
assert.equal(structuredRanked.candidates[0].match.queryCoverage, 100);
assert.equal(structuredRanked.candidates[0].match.missingOperations.length, 0);
assert.ok(
  structuredRanked.candidates[0].match.matched.some(
    (entry) => entry.field === "purpose" && entry.quality === 100,
  ),
);
const conflictingKickoff = structuredRanked.rejected.find(
  (entry) => entry.id === "artifact-template-project-kickoff",
);
assert.ok(conflictingKickoff);
assert.ok(conflictingKickoff.reasons.includes("avoid-when-conflict"));
assert.match(conflictingKickoff.conflicts[0].avoidWhen, /quarterly business review/i);

const noSemanticMatch = await queryTemplates({
  kind: "presentation",
  roots: [templateRoot],
  intent: {
    purposes: ["quantum entanglement laboratory protocol"],
  },
});
assert.equal(noSemanticMatch.retrievalStatus, "none");
assert.deepEqual(noSemanticMatch.candidates, []);
assert.equal(noSemanticMatch.rejected.length, expectedCounts.get("presentation"));
assert.ok(
  noSemanticMatch.rejected.every((entry) =>
    entry.reasons.includes("insufficient-relevance"),
  ),
);

const unsupportedOperation = await queryTemplates({
  kind: "presentation",
  roots: [templateRoot],
  intent: {
    purposes: ["quarterly business review"],
    requiredOperations: ["replace-embedded-workbook"],
  },
});
assert.equal(unsupportedOperation.retrievalStatus, "none");
assert.deepEqual(unsupportedOperation.candidates, []);
assert.ok(
  unsupportedOperation.rejected
    .find((entry) => entry.id === "artifact-template-business-review")
    .missingOperations.includes("replace-embedded-workbook"),
);

const brandSensitive = await queryTemplates({
  kind: "presentation",
  roots: [templateRoot],
  id: "artifact-template-business-review",
  intent: {
    purposes: ["quarterly business review"],
    brandSensitive: true,
  },
});
assert.ok(brandSensitive.candidates[0].reviewFlags.includes("brand-sensitive"));

const explicit = await queryTemplates({
  kind: "document",
  roots: [templateRoot],
  id: "artifact-template-system-design",
});
assert.deepEqual(explicit.candidates.map((candidate) => candidate.id), [
  "artifact-template-system-design",
]);
await assert.rejects(
  queryTemplates({
    kind: "spreadsheet",
    roots: [templateRoot],
    id: "artifact-template-system-design",
  }),
  /was not found for kind spreadsheet/,
);

const cli = spawnSync(
  process.execPath,
  [
    officeKitCli,
    "template",
    "search",
    "--kind",
    "spreadsheet",
    "--root",
    templateRoot,
    "--id",
    "artifact-template-sales-pipeline",
    "--json",
  ],
  { encoding: "utf8" },
);
assert.equal(cli.status, 0, cli.stderr);
const cliResult = JSON.parse(cli.stdout);
assert.equal(cliResult.selectionMade, false);
assert.equal(cliResult.candidates[0].id, "artifact-template-sales-pipeline");

const structuredCli = spawnSync(
  process.execPath,
  [
    officeKitCli,
    "template",
    "search",
    "--kind",
    "presentation",
    "--root",
    templateRoot,
    "--purpose",
    "quarterly business review",
    "--audience",
    "executive",
    "--content-shape",
    "KPIs",
    "--operation",
    "recognized-placeholder-title-text-replace",
    "--brand-sensitive",
    "--json",
  ],
  { encoding: "utf8" },
);
assert.equal(structuredCli.status, 0, structuredCli.stderr);
const structuredCliResult = JSON.parse(structuredCli.stdout);
assert.equal(structuredCliResult.ranking.algorithm, "bm25f");
assert.equal(
  structuredCliResult.candidates[0].id,
  "artifact-template-business-review",
);
assert.equal(structuredCliResult.queryIntent.brandSensitive, true);
assert.ok(
  structuredCliResult.candidates[0].reviewFlags.includes("brand-sensitive"),
);

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "office-kit-template-query-"));
try {
  await writeBrokenTemplate(tempRoot, "artifact-template-old-schema", {
    schemaVersion: 1,
    kind: "document",
  });
  await writeTemplate(tempRoot, "artifact-template-bad-hash", {
    provenance: {
      referenceSha256: "0".repeat(64),
    },
  });
  await writeTemplate(tempRoot, "artifact-template-path-escape", {
    reference: "../outside.docx",
  });
  await writeTemplate(tempRoot, "artifact-template-unknown-field", {
    undocumentedSelector: true,
  });
  await writeTemplate(tempRoot, "artifact-template-unknown-visual-trait", {
    visualTraits: {
      tone: [],
      density: "mixed",
      colorMode: "mixed",
      structure: [],
      undocumentedTrait: true,
    },
  });
  await writeTemplate(tempRoot, "artifact-template-non-english-search", {
    useWhen: ["季度业务复盘"],
  });
  await writeTemplate(tempRoot, "artifact-template-wrong-reference-kind", {
    reference: "assets/reference.xlsx",
  });
  await writeTemplate(tempRoot, "artifact-template-missing-skill");
  await fs.rm(path.join(tempRoot, "artifact-template-missing-skill", "SKILL.md"));
  const invalid = await queryTemplates({
    kind: "document",
    roots: [tempRoot],
    maxCandidates: 20,
  });
  assert.equal(invalid.candidates.length, 0);
  assert.equal(invalid.invalid.length, 8);
  assert.match(invalid.invalid.find((entry) => entry.id === "artifact-template-old-schema").error, /schemaVersion must be 2/);
  assert.match(invalid.invalid.find((entry) => entry.id === "artifact-template-bad-hash").error, /SHA-256 mismatch/);
  assert.match(invalid.invalid.find((entry) => entry.id === "artifact-template-path-escape").error, /safe relative path/);
  assert.match(
    invalid.invalid.find((entry) => entry.id === "artifact-template-unknown-field").error,
    /metadata contains unsupported fields: undocumentedSelector/,
  );
  assert.match(
    invalid.invalid.find((entry) => entry.id === "artifact-template-unknown-visual-trait").error,
    /visualTraits contains unsupported fields: undocumentedTrait/,
  );
  assert.match(
    invalid.invalid.find((entry) => entry.id === "artifact-template-non-english-search").error,
    /useWhen must use English search text/,
  );
  assert.match(
    invalid.invalid.find((entry) => entry.id === "artifact-template-wrong-reference-kind").error,
    /document templates must use a \.docx reference/,
  );
  assert.match(
    invalid.invalid.find((entry) => entry.id === "artifact-template-missing-skill").error,
    /SKILL\.md/,
  );
  await assert.rejects(
    queryTemplates({
      kind: "document",
      roots: [tempRoot],
      id: "artifact-template-bad-hash",
    }),
    /Requested template artifact-template-bad-hash is invalid/,
  );

  const lowerPriorityRoot = path.join(tempRoot, "lower-priority");
  await fs.mkdir(lowerPriorityRoot);
  await writeTemplate(lowerPriorityRoot, "artifact-template-bad-hash");
  const shadowed = await queryTemplates({
    kind: "document",
    roots: [tempRoot, lowerPriorityRoot],
    maxCandidates: 20,
  });
  assert.equal(
    shadowed.candidates.some((candidate) => candidate.id === "artifact-template-bad-hash"),
    false,
    "an invalid higher-priority template must not fall back to a lower-priority duplicate",
  );
  assert.ok(
    shadowed.invalid.some((entry) => entry.id === "artifact-template-bad-hash"),
    "the claimed invalid template must remain auditable",
  );
  const reversedPriority = await queryTemplates({
    kind: "document",
    roots: [lowerPriorityRoot, tempRoot],
    maxCandidates: 20,
  });
  assert.equal(
    reversedPriority.candidates.filter(
      (candidate) => candidate.id === "artifact-template-bad-hash",
    ).length,
    1,
    "the first valid catalog root must own a duplicate template id",
  );
  assert.equal(
    reversedPriority.invalid.some((entry) => entry.id === "artifact-template-bad-hash"),
    false,
  );
  await assert.rejects(
    queryTemplates({ kind: "document", roots: "not-an-array" }),
    /Template roots must be an array/,
  );
  await assert.rejects(
    queryTemplates({ kind: "document", roots: Array.from({ length: 21 }, () => tempRoot) }),
    /At most 20 template roots/,
  );
  await assert.rejects(
    queryTemplates({ kind: "document", roots: [tempRoot], tags: Array.from({ length: 21 }, () => "tag") }),
    /At most 20 query tags/,
  );
  await assert.rejects(
    queryTemplates({
      kind: "document",
      roots: [tempRoot],
      intent: { undocumentedIntent: true },
    }),
    /intent contains unsupported fields: undocumentedIntent/,
  );
  await assert.rejects(
    queryTemplates({
      kind: "document",
      roots: [tempRoot],
      intent: {
        purposes: Array.from({ length: 21 }, (_, index) => `purpose ${index}`),
      },
    }),
    /intent\.purposes must be an array of at most 20 strings/,
  );
  await assert.rejects(
    queryTemplates({
      kind: "document",
      roots: [tempRoot],
      intent: { brandSensitive: "yes" },
    }),
    /intent\.brandSensitive must be a boolean/,
  );

  const alias = path.join(path.dirname(tempRoot), `${path.basename(tempRoot)}-alias`);
  await fs.symlink(tempRoot, alias, "dir");
  try {
    await assert.rejects(
      queryTemplates({ kind: "document", roots: [alias] }),
      /Template root must be a real directory/,
    );
  } finally {
    await fs.rm(alias, { force: true });
  }
} finally {
  await fs.rm(tempRoot, { force: true, recursive: true });
}

const evalRecords = (await fs.readFile(path.join(repoRoot, "evals", "office-kit-routing.jsonl"), "utf8"))
  .split(/\r?\n/)
  .filter(Boolean)
  .map(JSON.parse);
assert.equal(evalRecords.length, 20);
assert.equal(new Set(evalRecords.map((record) => record.id)).size, 20);
assert.equal(evalRecords.filter((record) => record.expectedTrigger === "office-kit").length, 10);
assert.equal(evalRecords.filter((record) => record.expectedTrigger !== "office-kit").length, 10);
assert.ok(evalRecords.every((record) => record.prompt.length >= 80));
assert.ok(evalRecords.every((record) => ["consider", "explicit", "skip"].includes(record.templatePolicy)));

const templateEvalRecords = (
  await fs.readFile(
    path.join(repoRoot, "evals", "office-kit-template-selection.jsonl"),
    "utf8",
  )
)
  .split(/\r?\n/)
  .filter(Boolean)
  .map(JSON.parse);
assert.equal(templateEvalRecords.length, 14);
assert.equal(new Set(templateEvalRecords.map((record) => record.id)).size, 14);
assert.ok(templateEvalRecords.every((record) => record.prompt.length >= 60));
assert.ok(
  templateEvalRecords.every((record) =>
    ["clear", "unclear"].includes(record.goalClarity),
  ),
);
assert.ok(
  templateEvalRecords.every((record) =>
    ["unspecified", "named", "uploaded", "multiple-uploaded"].includes(
      record.templateChoice,
    ),
  ),
);
assert.ok(
  templateEvalRecords.every((record) =>
    ["search", "clarify", "template-elicit", "direct-feasibility", "register"].includes(
      record.expectedRoute,
    ),
  ),
);
assert.ok(
  templateEvalRecords.every((record) =>
    ["selected", "ask", "none", "defer", "register"].includes(
      record.expectedDecision,
    ),
  ),
);
assert.ok(
  templateEvalRecords
    .filter((record) => record.expectedRoute === "search")
    .every((record) => record.catalogQuery === true),
);
assert.ok(
  templateEvalRecords
    .filter((record) => record.expectedRoute !== "search")
    .every((record) => record.catalogQuery === false),
);
assert.ok(
  templateEvalRecords.some(
    (record) =>
      record.goalClarity === "unclear" &&
      record.templateChoice === "unspecified",
  ),
);
assert.ok(
  templateEvalRecords.some(
    (record) =>
      record.goalClarity === "unclear" &&
      ["named", "uploaded"].includes(record.templateChoice),
  ),
);
assert.ok(
  templateEvalRecords.filter((record) =>
    record.templateChoice.includes("uploaded"),
  ).length >= 6,
);

const skillValidatorPath = path.join(
  os.homedir(),
  ".codex",
  "skills",
  ".system",
  "skill-creator",
  "scripts",
  "quick_validate.py",
);
if (await exists(skillValidatorPath)) {
  const skillValidator = spawnSync("python3", [skillValidatorPath, skillRoot], {
    encoding: "utf8",
  });
  assert.equal(skillValidator.status, 0, `${skillValidator.stdout}\n${skillValidator.stderr}`);
}

const pluginValidatorPath = path.join(
  os.homedir(),
  ".codex",
  "skills",
  ".system",
  "plugin-creator",
  "scripts",
  "validate_plugin.py",
);
if (await exists(pluginValidatorPath)) {
  const pluginValidator = spawnSync("python3", [pluginValidatorPath, pluginRoot], {
    encoding: "utf8",
  });
  assert.equal(pluginValidator.status, 0, `${pluginValidator.stdout}\n${pluginValidator.stderr}`);
}

console.log("office-kit skill smoke ok");

async function readJson(filePath) {
  return JSON.parse(await fs.readFile(filePath, "utf8"));
}

async function exists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function writeBrokenTemplate(root, id, sidecar) {
  const template = path.join(root, id);
  await fs.mkdir(template, { recursive: true });
  await fs.writeFile(path.join(template, "artifact-template.json"), `${JSON.stringify(sidecar)}\n`);
}

async function writeTemplate(root, id, overrides = {}) {
  const template = path.join(root, id);
  const assets = path.join(template, "assets");
  const reference = Buffer.from("reference");
  const preview = Buffer.from("preview");
  await fs.mkdir(assets, { recursive: true });
  await Promise.all([
    fs.writeFile(
      path.join(template, "SKILL.md"),
      `---\nname: ${id}\ndescription: Test template.\n---\n`,
    ),
    fs.writeFile(path.join(assets, "reference.docx"), reference),
    fs.writeFile(path.join(assets, "preview.png"), preview),
  ]);
  const sidecar = {
    schemaVersion: 2,
    id,
    displayName: id,
    kind: "document",
    reference: "assets/reference.docx",
    preview: "assets/preview.png",
    useWhen: ["test"],
    avoidWhen: [],
    audiences: [],
    contentShapes: [],
    visualTraits: {
      tone: [],
      density: "mixed",
      colorMode: "mixed",
      structure: [],
    },
    visualCommitment: "opinionated",
    editProfile: {
      level: "copy-only",
      verifiedOperations: [],
    },
    provenance: {
      license: "user-provided",
      source: "test",
      referenceSha256: sha256(reference),
      previewSha256: sha256(preview),
      ...overrides.provenance,
    },
    ...Object.fromEntries(
      Object.entries(overrides).filter(([key]) => key !== "provenance"),
    ),
  };
  await fs.writeFile(path.join(template, "artifact-template.json"), `${JSON.stringify(sidecar)}\n`);
}
