## 1. Template contract

- [x] 1.1 Add schema-v3 presentation metadata validation and style-evidence search results while retaining schema-v2 DOCX/XLSX behavior.
- [x] 1.2 Reject source-backed presentation metadata and prohibited template assets with a specialist-creator migration error.
- [ ] 1.3 Update the existing template search contract tests with one v3 candidate and one v2 rejection.

## 2. Presentation Template Creator

- [x] 2.1 Add the `presentation-template-creator` plugin, concise Skill, format reference, Agent metadata, manifest, and deterministic packaging script.
- [x] 2.2 Implement safe create/update, PNG validation, montage generation, content hashes, provenance, and fixed-surface checks.
- [x] 2.3 Install and route the specialist by default; limit the generic creator to DOCX/XLSX.
- [ ] 2.4 Add one creator create/update contract flow to the existing template tests.

## 3. Presentation consumption

- [ ] 3.1 Converge OfficeKit and Presentations routing on zero-or-one style Skill selection, deck-specific Design Grammar, free Compose, and rendered review.
- [ ] 3.2 Separate templates from design systems, reference decks, and source-bound continuation in all shipped guidance.
- [ ] 3.3 Remove Grid defaults, fixed-layout fallback, source materialization, and old template-edit instructions from the presentation path.

## 4. Original bundled style Skills

- [x] 4.1 Rebuild Business Review as a schema-v3 style Skill with original guide and calibration images.
- [x] 4.2 Rebuild Market Trends Report as a schema-v3 style Skill with original guide and calibration images.
- [x] 4.3 Rebuild Operating Review as a schema-v3 style Skill with original guide and calibration images.
- [x] 4.4 Rebuild Project Kickoff as a schema-v3 style Skill with original guide and calibration images.
- [x] 4.5 Rebuild Simple Dark Mode as a schema-v3 style Skill with original guide and calibration images.
- [x] 4.6 Rebuild Simple Light Mode as a schema-v3 style Skill with original guide and calibration images.
- [x] 4.7 Rebuild Team Alignment as a schema-v3 style Skill with original guide and calibration images.
- [x] 4.8 Rebuild Grid Layout Library as a schema-v3 style Skill without layout code or skeletons.

## 5. Remove the old representation

- [ ] 5.1 Delete all bundled presentation reference PPTX files and legacy previews.
- [ ] 5.2 Delete embedded Grid modules, fixed layouts, registries, screenshots, support scripts, and fallback guidance.
- [ ] 5.3 Update template provenance, package inventory, third-party notices, and old round-trip claims without changing DOCX/XLSX templates.

## 6. Release evidence

- [ ] 6.1 Add package scans for the eight v3 style Skills and absence of PPTX/code/DSL template assets.
- [ ] 6.2 Run one fresh-context unrelated four-page creation through a selected style and record only concrete failures.
- [ ] 6.3 Update README, coverage, release metadata, and package version to 1.1.0.
- [ ] 6.4 Run narrow Skill/search/creator/package checks, then one final npm, package, and release verification pass.

## 7. Delivery

- [ ] 7.1 Commit OpenSpec, contracts, creator, routing, each rebuilt style, cleanup, and release evidence as focused atomic commits.
- [ ] 7.2 Push the feature branch normally and integrate main only after verifying current origin ancestry and a clean worktree.
