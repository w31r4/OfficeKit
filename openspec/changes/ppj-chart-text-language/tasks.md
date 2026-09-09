## 1. Field and lifecycle

- [x] 1.1 Add optional chart text language to schema/wire and reuse the language validator in shared native parsing/writing/semantics. Verify explicit presence, language-only styles, invalid tags and global-font-family profile rejection in focused tests.
- [x] 1.2 Propagate language through authored/source-bound conversion, projection, grammar precedence, rich styles, vector title defaults and vector label styles. Verify authored/source-bound add/change/delete/recreate, no-op bytes, non-target ZIP preservation and vector run override/label token with minimal fixtures.

## 2. Evidence and documentation

- [x] 2.1 Verify unrelated JS chart edits preserve wire language, and unknown native properties remain opaque. Run focused native and JS tests and proto:check with regenerated bindings.
- [x] 2.2 Update chart references, coverage and F-07 backlog with the bounded field evidence; regenerate derived documentation and run affected drift checks, strict OpenSpec validation and diff checks. Keep host proofreading and full F-07 open.

Evidence (2026-09-09): 19/19 focused native regressions passed with SDK 8.0.128 (PpjChartTextLanguage, PpjTrendlineRichText, PpjTrendlineLabel, PpjChartTextAlignment/Underline/Fill). Four language cases cover line/combo lifecycle across all fixture style owners, language-only presence, invalid authored/source-bound tokens and native tags, global-font-family rejection, semantic absence versus literal default-language, field precedence, and vector title/label propagation. The prior grammar sources/id fixture errors were corrected to the existing schema. JS worksheet chart wire preservation, proto:check after staging the generated binding, strict OpenSpec validation, maintainer check, capability-matrix drift, preview capability coverage, Skill portability (255 files), reference sync (333 files) and diff checks passed. No NativeAOT release rebuild or host proofreading/visual acceptance was performed. Full F-07 and the overall P0/P1 backlog remain open.
