## Context

The chart style parser owns a bounded DrawingML character profile shared by chart titles, legends, labels and axes. Trendline rich text reuses it for paragraph/run/end styles. Ordinary PPJ text already has strike boolean aliases and three native enum strings.

## Goals / Non-Goals

Complete direct strike state across existing chartTextStyle consumers. Preserve explicit cancellation separately from removing the property. Additional character effects, inherited strike editing and host visual evaluation remain separate work.

## Decisions

- Extract the existing textStyle strike vocabulary into a shared textStrike definition and reference it from both styles; the schema resolver supports direct definitions rather than nested property references. Reuse NativeStrike for booleans; fresh projection uses noStrike/sngStrike/dblStrike consistently.
- Add optional string strike=13 to SpreadsheetChartTextStyleArtifact. An optional field distinguishes omission from an explicitly empty invalid wire value, and preserves noStrike without treating it as missing.
- Extend shared native style parsing, writing, meaningful-style checks and semantic comparison; use exact native tokens, and keep the global-font-family-only guard.
- Propagate through authored and source-bound style mapping, nested grammar precedence, projection, rich-text styles and vector title/default label builders. Title run strike takes precedence over chart defaults.
- Reuse the language lifecycle fixture's style-owner helper for line/combo coverage instead of creating a separate parameter matrix. Use one vector fixture and a focused JS preservation assertion.
- Implement and test an isolated worktree, then stage its exact bytes so concurrent preview changes in shared compiler files cannot enter this commit.

## Risks / Trade-offs

- Explicit false could be mistaken for deletion → assert native noStrike and its reprojection separately from omission.
- An added field could accidentally relax unknown native topology → retain malformed strike and unknown-character-property opacity tests.
- Ordinary-text vocabulary has no grammar token form for strike → keep the same literal contract; nested field precedence remains supported.
