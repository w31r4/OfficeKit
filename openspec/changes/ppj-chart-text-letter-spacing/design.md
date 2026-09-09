## Context

Ordinary text already uses letterSpacing in points and writes spc by rounding points times 100. Chart text's shared codec rejects spc. The latest capitalization field established the same owner, precedence and vector propagation routes.

## Goals / Non-Goals

Preserve direct character spacing for existing chartTextStyle consumers. Font measurement, inherited style editing and renderer/host layout remain separate work.

## Decisions

- Extract ordinary text's numeric bounds into textLetterSpacing and reference it from both style definitions; keep the existing literal-only vocabulary.
- Add optional sint32 letter_spacing_hundredth_points=16, with -76800 through 76800 bounds. Native integers keep imported precision and explicit zero exact; a floating wire value would duplicate rounding decisions.
- Normalize PPJ points at the chart boundary using Math.Round with ties to even, consistent with ordinary text writing. Include spacing in validation, semantics, meaningful-style checks and the global-font-only restriction.
- Reuse shared native character property handling for chart and rich trendline styles. Propagate through authored/source-bound mapping, nested precedence and projection. Vector text maps to FontSpacingPoints, with explicit title-run values taking precedence.
- Reuse line/combo lifecycle fixtures and one vector fixture; verify the native spc attribute and literal text. Publish exact isolated bytes to preserve concurrent preview work.

## Risks / Trade-offs

- Zero mistaken for absence → presence-aware wire state and reset/removal regression.
- Precision differs between native and vector text → shared conversion plus hundredth-point assertions.
- Recognizing spc accidentally permits unknown styles → malformed spc and supported spc with unknown character properties remain opaque.
