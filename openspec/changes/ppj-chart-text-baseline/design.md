## Context

Ordinary PPJ text uses a finite baseline percentage in [-400,400] and writes native integer thousandths with Math.Round. Shared chart character styles currently omit baseline; their existing presence-aware language/strike fields provide the lifecycle pattern.

## Goals / Non-Goals

Preserve direct signed baseline state on existing chartTextStyle owners, including an explicit zero. Inherited layout, glyph shaping and host visual evaluation remain separate work.

## Decisions

- Extract the existing baseline numeric schema into textBaseline and reference it from ordinary and chart styles; keep its literal-number contract.
- Add optional sint32 baseline_thousandth_percent=14 to SpreadsheetChartTextStyleArtifact. Exact native integer state avoids floating-point semantic comparisons and retains explicit zero.
- Convert percentages once at the PPJ boundary with finite/range validation and Math.Round(value*1000, ToEven), matching ordinary text. Project native integers divided by 1000; document the 0.001% precision.
- Extend strict native parsing/writing, semantic and meaningful-style checks, and the global-font-family-only guard. Native integers outside [-400000,400000] or malformed values stay unsupported.
- Propagate through authored/source-bound mapping, grammar field precedence, rich text and vector titles/labels. Vector title runs keep their more-specific baseline.
- Reuse the existing shared line/combo style-owner fixture and one vector fixture, then publish exact isolated bytes to preserve concurrent preview work.

## Risks / Trade-offs

- Zero might disappear as a default → use optional native state and test zero separately from deletion.
- Percent units could be confused with native thousandths → assert exact native positive/negative/zero values and non-integral conversion.
- Recognizing baseline could unlock unsupported character effects → keep malformed baseline and unknown cap properties opaque.
