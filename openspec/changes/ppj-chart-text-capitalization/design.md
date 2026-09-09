## Context

Ordinary text uses capitalization none/small/all and the native cap attribute. The shared ChartML style codec already preserves optional language, strike and baseline but still rejects cap.

## Goals / Non-Goals

Preserve direct capitalization display state for existing chartTextStyle consumers. Literal text remains unchanged. Inherited style editing, glyph selection and host appearance remain separate work.

## Decisions

- Extract the existing enum into textCapitalization and reference it from ordinary and chart styles, retaining the same literal-only vocabulary.
- Add optional string capitalization=15 to SpreadsheetChartTextStyleArtifact, preserving explicit none and rejecting empty/unknown values.
- Extend shared native style parsing, writing, semantic comparison, meaningful-style checks and global-font-family restriction. A capitalization-only style is valid.
- Propagate through authored/source-bound style mapping, field precedence, rich paragraph/run/end styles and projection. Vector text maps to FontCaps, and title defaults yield to explicit run capitalization.
- Reuse the shared line/combo style-owner fixture and one vector fixture. Compare unchanged native text as well as style state. Publish exact isolated bytes so concurrent preview changes stay outside the commit.

## Risks / Trade-offs

- Treating none as absent would lose a direct override → verify none and omission separately.
- Applying string case conversion would destroy literal data → preserve original chart and run characters throughout the lifecycle.
- Recognizing cap could relax unknown character graphs → retain malformed cap and unsupported character-property opacity tests.
