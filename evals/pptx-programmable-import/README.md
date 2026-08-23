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

## Frozen baseline result

The committed baseline evaluates product commit
`d5df8df94727dccd4412e6be874d1c5407b57f64` through packed
`office-kit@0.6.0` tarball SHA-256
`2bec8a4caf4c15f840be4005424111a4a6207b49c24fe23c25908838b4095120`.
It is intentionally a failed acceptance baseline:

- The deterministic matrix completed all 90 runs. 60 runs passed source,
  package, relationship, masked XML/SVG, second-import, and pixel checks; 10 of
  30 intents were byte-deterministic across all three runs.
- 算秩未来 passed 18/30 runs and 6/10 deterministic intents. Four text edits
  changed the package but produced no target-page pixel change.
- 蓝灰酸性模板 passed 12/30 runs and 4/10 deterministic intents. Six text or
  position edits changed the package but produced no target-page pixel change.
- 麦肯锡 SVG passed all 30 per-run package/pixel checks, but 0/10 intents were
  byte-deterministic because each repetition issued a different copy-on-write
  relationship ID. Those IDs and output hashes remain unnormalized in the
  evidence.
- All nine fresh-context Codex trials used isolated packed installs, preserved
  the read-only source hash, and passed the forbidden-path scan. None published
  an output: four stopped on JSONL/JavaScript authoring errors, two omitted the
  required commit summary, and three matched the SVG node through the wrong
  record field. Their complete bounded final explanations remain in evidence.

The immutable components are `baseline/matrix.v1.json` and
`baseline/codex.v1.json`; `baseline.v1.json` records their hashes and derives
the summary. The matrix component also records an evaluator-only replay over
the retained 90 outputs: no edit was rerun, no outcome changed, and package,
relationship, masked-leaf, second-import, and pixel details remain present even
when the final pixel check failed. No absent Codex output is credited with
package, pixel, second-import, or task/resume success.
