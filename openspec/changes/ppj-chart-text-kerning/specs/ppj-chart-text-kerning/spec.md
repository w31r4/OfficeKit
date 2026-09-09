## Purpose

Express direct chart kerning thresholds in points while preserving native precision, optional presence and the source-bound editing lifecycle.

## ADDED Requirements

### Requirement: Kerning threshold and presence

Existing chartTextStyle consumers SHALL accept finite kerning values from 0 through 768 points, sharing ordinary text syntax. Values SHALL round to native hundredths of a point with ties to even. Explicit zero SHALL remain distinct from omission, and a kerning-only style SHALL be meaningful. Literal characters SHALL remain unchanged and the global-font-family-only profile SHALL remain restricted.

#### Scenario: Change threshold then remove it
- **WHEN** a chart style changes from a positive threshold to zero, then omission and recreation
- **THEN** native kern and fresh projection SHALL preserve the corresponding value or absence at 0.01pt precision

### Requirement: Existing style owners retain kerning

Title, legend, axes, data-label and trendline styles, including rich paragraph/run/end styles, SHALL support authored and source-bound creation, change, zero, removal and recreation. Declared field precedence SHALL apply. Vector labels SHALL retain kerning and explicit title-run kerning SHALL override chart defaults. No-op SHALL preserve original bytes, and native chart style changes SHALL preserve non-target ZIP entries.

#### Scenario: Edit and reproject native charts
- **WHEN** line or combo chart kerning is edited from a fresh native projection
- **THEN** only the target ChartPart SHALL change and fresh projection SHALL recover the requested normalized field or absence

### Requirement: Invalid and unknown state remains protected

Negative, out-of-range, non-finite or malformed kerning SHALL fail closed. Unknown imported character properties SHALL remain source-owned. Unrelated JS chart edits SHALL retain optional kerning including zero.

#### Scenario: Unsupported imported character properties
- **WHEN** an owner contains malformed kern or another unsupported character property
- **THEN** unsupported edits SHALL remain unavailable and no-op SHALL preserve source bytes
