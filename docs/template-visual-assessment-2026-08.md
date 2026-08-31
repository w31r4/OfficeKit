# Presentation template visual assessment

> Date: 2026-08-31  ·  Scope: 30 Kimi-derived + 8 Codex-aligned presentation templates (38 total). Evidence Ledger is tracked separately and is not counted in this requested set.

## Result

All 38 templates reached the declared threshold: visual index 100–100 and functional index 100–100. The run rendered 323 pages through LibreOffice + Poppler after compiling each reference PPJ once.

“Visual index” here means reference-package render parity: the compiled PPJ output SHA-256 equals the sidecar reference PPTX SHA-256, and every compiled page produced a non-empty PNG with valid dimensions. It is deliberately not a subjective claim that unrelated clean-room content is pixel-identical to the original Kimi inspiration images. A human style review remains a separate activity.

## Method

- `officekit ppj check` was run once for each template; all 38 PPJ dependency closures passed.
- Each reference PPJ was compiled once through the packaged NativeAOT codec and every page was rendered with LibreOffice + Poppler.
- The compiled PPTX hash was compared with the declared `referencePptx` hash; every page image was checked for valid PNG dimensions and non-empty bytes.
- Functional score covers schema/dependency validation, deterministic reference package identity, and the declared source-bound/authored boundary. It does not claim arbitrary third-party edits are supported.
- Evaluation artifacts are kept in disposable `/tmp` output; this report stores hashes and counts only.

## Per-template evidence

| Template | Origin | Pages | Rendered | Examples / roles | Source bound | Visual | Functional | Status |
| --- | --- | ---: | ---: | ---: | :---: | ---: | ---: | --- |
| artifact-template-amber-committee-memo | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-apricot-dossier | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-aqua-impact-story | kimi | 6 | 6 | 6 / 6 | no | 100 | 100 | restored |
| artifact-template-axis-atlas | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-blue-flame-operations | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-blueprint-lecture | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-business-review | codex | 14 | 14 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-clay-craft-review | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-coastal-analysis | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-coral-growth-brief | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-cranberry-evidence | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-cream-civic-collage | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-ebony-investment-review | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-forest-strategy | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-gilt-market-ledger | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-grid-layout-library | codex | 26 | 26 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-indigo-verdict | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-jade-annual-brief | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-lake-research-journal | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-market-trends-report | codex | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-midnight-prospectus | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-moonlit-work-report | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-moss-transformation | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-noir-field-pictorial | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-operating-review | codex | 31 | 31 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-paper-seminar | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-project-kickoff | codex | 12 | 12 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-rice-paper-yearbook | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-river-handbook | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-saffron-editorial | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-silver-atelier | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-simple-dark-mode | codex | 4 | 4 | 4 / 4 | no | 100 | 100 | restored |
| artifact-template-simple-light-mode | codex | 26 | 26 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-skyline-wayfinding | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-soft-proof | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-team-alignment | codex | 24 | 24 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-tidal-research | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |
| artifact-template-violet-operations | kimi | 6 | 6 | 4 / 4 | yes | 100 | 100 | restored |

Full machine-readable evidence is in [`template-visual-assessment-2026-08.json`](./template-visual-assessment-2026-08.json).
