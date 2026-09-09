## Purpose

Express direct chart character spacing in points while preserving native precision, optional presence and the source-bound editing lifecycle.

## ADDED Requirements

### Requirement: Signed spacing and direct presence

Existing chartTextStyle consumers SHALL accept finite letterSpacing values from -768 through 768 points, sharing ordinary PPJ text syntax. Values SHALL round to native hundredths of a point with ties to even. Explicit zero SHALL remain distinct from omission, a spacing-only style SHALL be meaningful, and literal characters SHALL remain unchanged. The global-font-family-only profile SHALL remain restricted.

#### Scenario: Reset and remove spacing
- **WHEN** a chart style changes from positive to negative spacing, then zero, then omission
- **THEN** native spc and fresh projection SHALL preserve those states at 0.01pt precision without modifying the text

### Requirement: Spacing propagates through existing style owners

Title, legend, axes, data-label and trendline styles, including rich paragraph/run/end styles, SHALL support authored and source-bound creation, change, reset, removal and recreation. Declared field precedence SHALL apply. Vector labels SHALL retain spacing and explicit title-run spacing SHALL override chart defaults. No-op SHALL preserve original bytes; native chart style changes SHALL preserve non-target ZIP entries.

#### Scenario: Source edit followed by fresh projection
- **WHEN** line or combo chart spacing is edited from a fresh native projection
- **THEN** only the target ChartPart SHALL change and fresh PPJ projection SHALL reflect the requested normalized value or absence

### Requirement: Invalid native state stays protected

Out-of-range, non-finite or malformed spacing SHALL fail closed. Unknown imported character properties SHALL remain source-owned. Unrelated JS chart edits SHALL retain optional signed spacing including explicit zero.

#### Scenario: Unsupported imported character graph
- **WHEN** a chart owner has malformed spc or another unsupported character property
- **THEN** unsupported style edits SHALL remain unavailable and no-op SHALL preserve the original source
