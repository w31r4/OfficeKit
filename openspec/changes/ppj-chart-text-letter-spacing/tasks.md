## 1. Spacing field and lifecycle

- [x] 1.1 Add shared numeric schema, optional integer wire state and native spacing parsing/writing/semantics. Verify signed bounds, rounding, explicit zero and global-font-only protection.
- [x] 1.2 Propagate through authored/source-bound mapping, projection, rich styles, field precedence and vector text. Verify line/combo creation/change/reset/deletion/recreation, literal text, exact no-op/non-target ZIP preservation and vector overrides.

## 2. Evidence and documentation

- [x] 2.1 Keep malformed spc/unknown character state opaque and retain signed/zero wire values through unrelated JS edits. Run focused native/JS regressions and proto:check on the isolated snapshot.
- [x] 2.2 Update chart guidance, registry, coverage and F-07 backlog; regenerate manual/matrix and pass affected drift, portability, reference sync, strict OpenSpec and diff checks. Record host/runtime evidence limits.

Evidence (2026-09-09): the isolated snapshot based on `6c4c841e` passed 35 selected native regressions with SDK 8.0.128, including four letter-spacing cases and adjacent chart text/trendline cases. Line/combo fixtures verify every shared chart/rich-label style owner through creation, signed changes, zero, deletion, recreation, exact no-op, target-only ChartPart changes, non-target ZIP byte preservation, unchanged text and fresh projection. Other cases verify bounds/non-finite inputs, native hundredth-point conversion and ties, explicit-zero semantics, global-font restriction, field precedence and vector title/label spacing. Malformed spc and supported spc with unsupported kern remain opaque with exact no-op and refused analytics edits.

JS worksheet preservation and optional signed/zero protobuf round-trips passed. `npm run proto:check` passed after staging only the intentional generated binding in this isolated worktree. Generated manual/matrix checks, preview capability coverage, Skill portability (255 files), reference-source sync (333 files in the main checkout), strict OpenSpec validation and whitespace checks passed. No NativeAOT rebuild, glyph measurements or PowerPoint host acceptance was performed. Full F-07 and the P0/P1 backlog remain open.
