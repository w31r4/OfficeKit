## Purpose

Preserve independently editable inner-shadow state on chart character styles through authored output, native import and source-bound edits.

## ADDED Requirements

### Requirement: Typed chart character inner shadow
Chart text styles SHALL accept innerShadow with required RGB/theme/grammar color and optional blur, distance, angle and opacity. Bounds SHALL be 0..1000 points for blur, 0..100000 points for distance, -360..360 degrees for angle and 0..1 for opacity. Omission SHALL remain distinct from explicit zero, including wire transport. Named field precedence and explicit vector run overrides SHALL apply.

#### Scenario: Presence and precedence
- **WHEN** a chart uses an inner-shadow-only style with absent geometry or explicit zero and named or inline overrides
- **THEN** the style is meaningful, native presence is retained on projection, and configured field precedence and explicit run overrides determine output

### Requirement: Independent native effect lifecycle
Recognized chart character owners SHALL preserve glow, inner shadow, outer shadow and soft edge in that order. Creation, change, deletion and recreation of innerShadow SHALL preserve unrelated effects, literal text and non-target ZIP parts. Native no-op compilation SHALL retain exact source bytes.

#### Scenario: Line and combo lifecycle
- **WHEN** a fresh projected line or combo chart has its innerShadow changed or deleted across title, legend, axes, data labels and rich trendline label styles
- **THEN** only its chart part changes, fresh projection reflects the selected inner shadow and unrelated effects/text remain intact

### Requirement: Unsupported native graphs retain source ownership
Malformed, duplicate, reordered, unknown or extended effect graphs and out-of-bound inner-shadow values SHALL be rejected for typed editing while source-bound no-op preserves original bytes.

#### Scenario: Invalid effect owner
- **WHEN** an imported effect owner contains duplicate inner shadows, nested alpha extensions, extra attributes or excessive geometry
- **THEN** it does not gain the chart analytics edit capability and attempts to edit it fail closed
