# PPTX programmable-import acceptance

This directory defines an evaluator-owned acceptance platform for the three
frozen real PPTX samples. It tests the published package and Presentations
Skill; it is not runtime evidence by itself.

- `intent-matrix.v1.json` contains 30 bounded edit intents. Every intent starts
  from the immutable source and must produce the same bytes and oracle result
  in three independent runs.
- `continuation-tasks.v1.json` contains three complete task/resume/publish
  tasks. Each task is executed in three fresh ephemeral Codex contexts.
- `baseline.v1.json` is generated only after the packed clean-install matrix
  and all nine Codex trials have been evaluated. Failures stay failures.

The Agent may use only the installed OfficeKit Presentations Skill and public
`office-kit` API. Raw package/XML editing, Python, HTML/PPTD,
`@oai/artifact-tool`, blank-deck reconstruction, source overwrite, and silent
fallback are disqualifying. Independent evaluator code owns OPC, relationship,
masked XML/SVG, second-import, source-hash, non-target-pixel, task/resume, and
no-overwrite checks.

Run the deterministic matrix from a create-only directory:

```sh
node scripts/pptx-programmable-import-matrix.mjs \
  --assets-dir /absolute/path/to/frozen-assets \
  --package-root /absolute/path/to/packed-clean-install/node_modules/office-kit \
  --install-kind packed-clean-install \
  --tarball-sha256 <sha256> \
  --run-root /absolute/create-only/path
```

Run the nine fresh-context continuation trials independently:

```sh
node scripts/pptx-programmable-import-codex-harness.mjs \
  --assets-dir /absolute/path/to/frozen-assets \
  --run-root /absolute/create-only/path
```

The Codex harness packs the candidate once, installs that tarball into every
isolated trial with lifecycle scripts disabled, installs only the public
OfficeKit and Presentations Skills, and retains the full trace, stderr, task
store, candidate revisions, renderer cache, and evaluator report outside the
repository. A non-zero exit still writes `evidence.json`; it never converts a
trial failure into a skip or edits the product under test.
