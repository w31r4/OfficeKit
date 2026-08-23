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
