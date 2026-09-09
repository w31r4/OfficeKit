## Context

Ordinary PPJ text already exposes kerning in points and writes native kern using points times 100. The chart style codec rejects kern. Existing letter-spacing lifecycle and shared-owner fixtures provide the propagation routes.

## Goals / Non-Goals

Preserve direct kerning thresholds on existing chartTextStyle consumers. Font pair metrics, host layout and inherited-style editing remain separate work.

## Decisions

- Extract ordinary text's finite 0..768 bounds into textKerning, retaining literal numeric syntax for both owners.
- Add optional uint32 kerning_hundredth_points=17, bounded to 76800. Native nonnegative integer state preserves exact imported precision and explicit zero; normalize PPJ values using ties-to-even rounding as ordinary text does.
- Extend shared native character parsing/writing, semantics, meaningful-style checks and the global-font-only restriction. Propagate through authored/source-bound mapping, projection, rich styles and nested precedence.
- Map vector text to FontKerningPoints; explicit title-run values override chart defaults. Preserve literal characters.
- Reuse line/combo and vector fixtures. Update the unsupported graph fixture from valid kern to valid kern plus unsupported altLang, and add malformed kern coverage.
- Kerning is a minimum font-size threshold, as described in Microsoft's DrawingML primer (https://download.microsoft.com/download/e/1/4/e14fb96f-83b8-4a2a-84db-7fa8acbe061a/Office%20Open%20XML%20Part%203%20-%20Primer.pdf). Do not document zero as disabling kerning or reuse signed letter-spacing semantics.

## Risks / Trade-offs

- Zero mistaken for absence → optional wire state and lifecycle checks.
- Negative spacing accidentally accepted as kerning → separate nonnegative bounds at every boundary.
- Recognizing kern relaxes unknown graphs → malformed kern and unsupported character-property regressions retain no-op and refusal behavior.
