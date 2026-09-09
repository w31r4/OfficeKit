## 1. Kerning field and lifecycle

- [x] 1.1 Add shared numeric schema, optional integer wire state and native kerning parsing/writing/semantics. Verify nonnegative bounds, rounding, zero presence and global-font-only protection.
- [x] 1.2 Propagate through authored/source-bound mapping, projection, rich styles, field precedence and vector text. Verify line/combo creation/change/zero/deletion/recreation, unchanged text, exact no-op/non-target ZIP bytes and vector overrides.

## 2. Evidence and documentation

- [x] 2.1 Keep malformed kern/unknown character state opaque and retain optional/zero wire values through unrelated JS edits. Run focused native/JS regressions and proto:check on the isolated snapshot.
- [x] 2.2 Update chart guidance, registry, coverage and F-07 backlog; regenerate manual/matrix and pass affected drift, portability, reference sync, strict OpenSpec and diff checks. Record threshold semantics and host/runtime evidence limits.

Evidence (2026-09-09): the isolated snapshot based on `f3ebbb34` passed 39 selected native regressions with SDK 8.0.128, including four kerning cases and adjacent chart text/trendline cases. Line/combo fixtures verify every shared chart/rich-label style owner through creation, change, explicit zero, deletion, recreation, exact no-op, target-only ChartPart changes, non-target ZIP byte preservation, unchanged literal text and fresh projection. Other cases verify nonnegative bounds/non-finite inputs, native hundredth-point precision and ties, zero semantics, global-font restriction, field precedence and vector title/label thresholds. Malformed kern and supported kern with unsupported altLang remain opaque with exact no-op and refused analytics edits.

JS worksheet preservation and optional/zero protobuf round-trips passed. `npm run proto:check` passed after staging only the intentional generated binding in this isolated worktree. Generated manual/matrix checks, preview capability coverage, Skill portability (255 files), reference-source sync (333 files in the main checkout), strict OpenSpec validation and whitespace checks passed. The downloaded Microsoft DrawingML primer, PDF page 299, confirms the minimum-font-size meaning and default-zero threshold; browser PDF parsing was unavailable, so the downloaded document was read with pypdf. No NativeAOT rebuild, font pair measurements or PowerPoint host acceptance was performed. Full F-07 and the P0/P1 backlog remain open.
