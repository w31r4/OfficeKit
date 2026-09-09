## Context

Ordinary text resolves highlight through Catalog.Color and rejects alpha below one. Chart styles already preserve fill and typography, but the shared character parser only consumes paint followed by typefaces. Direct a:highlight belongs between these retained child groups.

## Goals / Non-Goals

Preserve direct opaque RGB highlighting for existing chartTextStyle owners, using established PPJ colors. Theme-bound native highlight, color transforms, inherited state and host glyph appearance remain separate work.

## Decisions

- Reuse the existing color schema; add optional string highlight_rgb=18 to the chart text style wire message. Validate exact six-digit RGB, reject empty values and canonicalize casing. The ordinary authored highlight contract already resolves tokens to RGB, so no second theme/alpha model is introduced here.
- Parse exactly one attribute-free highlight with one attribute-only srgbClr child and reject extra text/children/attributes. Write it after text fill and before Latin/East Asian/complex-script typefaces. Verify this order using the Open XML character-property validator.
- Include highlight presence and canonical RGB in meaningful-style checks, semantics and the global-font-only restriction. Foreground fill and highlight remain independent.
- Resolve color/tint/shade in authored and source-bound paths; reject nonopaque results through a shared helper. Propagate nested field precedence, rich styles and projection. Vector text maps to HighlightRgb; title defaults yield to explicit run highlight state.
- Reuse line/combo fixtures for native lifecycle, token resolution, literal text and ZIP scope. A compact mixed-style case verifies foreground/highlight/typeface coexistence, invalid/opaque behavior and vector overrides. Publish isolated tested bytes.

## Risks / Trade-offs

- Child ordering corrupts native character properties → standalone Open XML validation of mixed properties and source lifecycle assertions.
- Highlight accidentally replaces foreground color → independent wire property and coexistence regression.
- Tokens silently drop alpha or native transforms → reject unresolved/nonopaque requests and unsupported imported graphs.
