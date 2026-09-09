## 1. Capitalization field and lifecycle

- [x] 1.1 Add shared enum schema, optional wire state and native style parsing/writing/semantics. Verify explicit none, invalid values and global-font-only protection.
- [x] 1.2 Propagate capitalization through authored/source-bound mapping, projection, rich styles, field precedence and vector text. Verify line/combo creation/change/reset/deletion/recreation, unchanged characters, exact no-op/non-target ZIP preservation and one vector override fixture.

## 2. Evidence and documentation

- [x] 2.1 Keep malformed cap/unknown native character state opaque and preserve wire capitalization through unrelated JS edits. Run focused native/JS regressions and proto:check on the isolated snapshot.
- [x] 2.2 Update chart guidance, registry, coverage and F-07 backlog; regenerate manual/matrix and pass affected drift, portability, strict OpenSpec and diff checks. Record host/runtime evidence limits.

Evidence (2026-09-09): implementation and verification used an isolated worktree based on `ae8f9404`. SDK 8.0.128 passed all 31 selected native regressions, including four capitalization cases and adjacent chart text/trendline cases. The line/combo lifecycle checks all shared style owners, exact no-op, target-only ChartPart edits, non-target ZIP bytes, unchanged literal text and fresh projection. Separate cases cover invalid wire/authored/native values, explicit none versus absence, global-font protection, nested precedence and vector title defaults/overrides and labels. Malformed cap and supported cap combined with unsupported kern remain opaque with exact no-op and rejected analytics edits. JS worksheet preservation and optional protobuf round-trips passed.

`npm run proto:check` passed after staging only the intentional generated binding in this isolated worktree. Presentation manual and matrix generation/checks, preview capability coverage, Skill portability (255 files), strict OpenSpec validation and whitespace checks passed; the main checkout's reference-source sync gate passed (333 files). No NativeAOT rebuild, font glyph measurement or PowerPoint host acceptance was performed. Full F-07 and the P0/P1 backlog remain open.
