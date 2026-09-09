## 1. Highlight field and lifecycle

- [x] 1.1 Add schema/wire/native RGB highlight state, validation, semantics and presence. Verify mixed paint/highlight/typeface order with Open XML validation, standalone highlight and global-font-only protection.
- [x] 1.2 Propagate colors/tokens through authored/source-bound mapping, rich styles, projection, precedence and vector text. Verify line/combo creation/change/removal/recreation, unchanged text/non-target ZIP bytes, exact no-op, token transforms and vector overrides.

## 2. Evidence and documentation

- [x] 2.1 Reject invalid/nonopaque colors and preserve malformed/duplicate/theme/transform/unknown native graphs; retain wire highlight in unrelated JS edits. Run focused native/JS regressions and proto:check in the isolated snapshot.
- [x] 2.2 Update chart guidance, registry, coverage and F-07 backlog; regenerate manual/matrix and pass affected drift, portability, reference sync, strict OpenSpec and diff checks. Record host/runtime limits.

Evidence (2026-09-09): the isolated snapshot based on `4782e6e5` passed 43 selected native regressions with SDK 8.0.128, including four highlight cases and adjacent chart text/trendline cases. Line/combo fixtures verify shared chart/rich-label owners through creation, color/token changes, tint/shade resolution, deletion, recreation, exact no-op, target-only ChartPart changes, non-target ZIP bytes, unchanged text and fresh projection. Title foreground color and typeface survive every highlight edit. A standalone mixed character-property case passes Office 2021 Open XML validation and verifies fill/highlight/typeface order; other cases cover highlight-only styles, case normalization, global-font restriction, invalid tokens/alpha on authored and source-bound paths, nested precedence and vector title/label overrides. Empty, duplicate, theme, transformed and malformed native highlight graphs remain source-owned, with exact no-op and refused analytics edits.

JS worksheet preservation and optional RGB protobuf round-trips passed. `npm run proto:check` passed after staging only the intentional generated binding in this isolated worktree. Generated manual/matrix checks, preview capability coverage, Skill portability (255 files), reference-source sync (333 files in the main checkout), strict OpenSpec validation and whitespace checks passed. No NativeAOT rebuild, renderer or PowerPoint host appearance acceptance was performed. Full F-07 and the P0/P1 backlog remain open.
