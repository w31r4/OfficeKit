---
name: skill-update
description: Check which OfficeKit Skills, Help records, API docs, examples, tests, and release evidence must be reviewed when a runtime, protocol, or workflow capability changes. Use for capability changes and Skill maintenance; it is read-only and never edits the repository.
---

# Skill Update

Use this maintenance Skill after changing an OfficeKit primitive, protocol
field, Help entry, public example, review invariant, or Skill route. It keeps
the ownership graph visible without forcing every Agent to read the whole
repository.

## Fast path

From the repository root (or the installed Skill directory):

```bash
node "$SKILL_DIR/scripts/check-primitive-impact.mjs" check --repo .
node "$SKILL_DIR/scripts/check-primitive-impact.mjs" impact --repo .
```

Use `--json` for a machine-readable report. `impact` reads the current Git
diff and untracked paths by default; pass `--paths file1 file2` when reviewing
a deliberately bounded change. The checker only reads text, JSON and Git
metadata. It does not run OfficeKit, initialize a provider, fetch a URL, or
write a file.

## What the report means

- **family**: the semantic capability area affected by the path;
- **help**: the Help names that must stay discoverable;
- **consumers**: the route, task, reference, or Creator Skill that teaches the
  capability;
- **examples/tests**: the smallest concrete workflow and focused protection;
- **evidence**: API docs, coverage, or release material that may need update.

The presentation primitive map lives at
`skills/presentations/skills/presentations/references/primitive-impact.json`.
It is the maintenance contract, not an API implementation. If a capability
does not fit an existing family, add a family and its consumer surfaces before
shipping the capability.

## Change protocol

1. Change the authoritative runtime or protocol source.
2. Update the Help catalog; regenerate `docs/api.md` when Help changes.
3. Run `impact` and inspect every reported consumer. Update only the Skills and
   references that own the behavior; do not duplicate the same prose elsewhere.
4. Add or adjust one focused example or assertion when the change is a new
   contract or a reproducible failure. Do not create a matrix solely to raise
   coverage.
5. Run `check` and the narrowest affected gate. Record remaining host or visual
   limits in the task/coverage evidence.

## Boundaries

This Skill does not replace Presentations, Documents, Spreadsheets, PDF, or
Template Creator. It does not author a second DSL, inspect private provider
bundles, infer legal compatibility, or make a visual quality claim. It reports
what must be considered; the owning Skill decides how to teach and review the
capability.
