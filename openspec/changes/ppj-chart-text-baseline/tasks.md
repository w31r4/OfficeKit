## 1. Baseline field and lifecycle

- [x] 1.1 Add shared percentage schema, optional signed native-unit wire state and native style parsing/writing/semantics. Verify precision, explicit zero, range/non-finite rejection and global-font-only protection.
- [x] 1.2 Propagate baseline through authored/source-bound mapping, projection, rich styles, field precedence and vector text. Verify line/combo creation/change/reset/deletion/recreation, exact no-op/non-target ZIP preservation and one vector override fixture.

## 2. Evidence and documentation

- [x] 2.1 Retain opacity for invalid/unknown native character state and preserve signed/zero wire baseline through unrelated JS edits. Run focused native/JS regressions and proto:check on the isolated snapshot.
- [x] 2.2 Update chart guidance, registry, coverage and F-07 backlog; regenerate manual/matrix and pass affected drift, portability, strict OpenSpec and diff checks. Record native precision and host/runtime evidence limits.

Evidence (2026-09-09): isolated snapshot based on e092c458 passed 27/27 focused native regressions using SDK 8.0.128 (PpjChartTextBaseline/Strike/Language, PpjTrendlineRichText/Label, PpjChartTextAlignment/Underline/Fill). Four baseline cases cover line/combo lifecycle and all fixture style owners, native signed units, zero versus omission, fractional rounding, exact no-op/non-target ZIP preservation, range/non-finite rejection, ties-to-even, global-font restriction, field precedence and vector title/label output. Unknown character attributes and malformed native baseline keep analytics closed. JS signed/zero wire preservation, proto:check with regenerated/staged binding, strict OpenSpec, maintainer, capability-matrix drift, preview capability coverage, Skill portability (255 files) and diff checks passed in the isolated worktree; reference sync (333 files) passed in the main worktree with its local legacy inputs. No NativeAOT release build or host visual acceptance was performed. Full F-07 and the P0/P1 backlog remain open.
