## Purpose

Express direct chart text highlighting through the existing PPJ color vocabulary while preserving text, other typography and source-bound edit scope.

## ADDED Requirements

### Requirement: Highlight paint is independent character state

Existing chartTextStyle consumers SHALL accept highlight using ordinary PPJ color syntax and declared color tokens, including tint/shade resolution. Fully opaque colors SHALL lower to direct RGB highlight. Projection SHALL emit canonical RGB. A highlight-only style SHALL be meaningful, highlight SHALL coexist with text fill and font properties, and literal characters SHALL remain unchanged. The global-font-family-only profile SHALL remain restricted.

#### Scenario: Highlight coexists with text paint and fonts
- **WHEN** a chart text style declares highlight, foreground color and a font family
- **THEN** native character properties SHALL retain all three in valid native order and fresh projection SHALL preserve the independent values

### Requirement: Highlight lifecycle propagates across style owners

Title, legend, axes, data-label and trendline styles, including rich paragraph/run/end styles, SHALL support authored and source-bound creation, change, removal and recreation. Declared field precedence SHALL apply. Vector labels SHALL retain highlight and explicit title-run highlight SHALL override chart defaults. No-op SHALL preserve original bytes and native chart style changes SHALL preserve non-target ZIP entries.

#### Scenario: Change and remove highlight
- **WHEN** line or combo chart highlight is changed, removed and recreated from fresh native projections
- **THEN** only the target ChartPart SHALL change and fresh PPJ projection SHALL recover the requested RGB value or absence

### Requirement: Unsupported highlight state stays source-owned

Invalid colors/tokens, nonopaque highlight requests, malformed or duplicate native highlight containers, native color transforms, theme highlight and unknown character properties SHALL fail closed. Unrelated JS chart edits SHALL retain optional highlight RGB state.

#### Scenario: Unknown native highlight graph
- **WHEN** an imported owner contains highlight with unsupported native color state
- **THEN** unsupported style edits SHALL remain unavailable and no-op SHALL preserve the source
