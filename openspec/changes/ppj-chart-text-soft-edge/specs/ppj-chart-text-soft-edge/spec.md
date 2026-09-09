## Purpose

Expose chart text soft-edge radius and preserve each known character effect independently through authoring and source-bound edits.

## ADDED Requirements

### Requirement: Chart text soft-edge field
Chart text styles SHALL accept `softEdge` with required numeric `radius` from 0 to 1000 points, using native EMU precision with ties-to-even rounding. Explicit zero SHALL remain distinct from omission. The field SHALL propagate to chart title, legend, axis, data-label and trendline label paragraph/run/end styles.

#### Scenario: Radius and zero presence
- **WHEN** authored or source-bound chart styles set radius to zero or a fractional point value
- **THEN** native character soft-edge radius and fresh PPJ projection SHALL preserve the effective value and field presence.

### Requirement: Independent known effect lifecycle
Soft edge SHALL coexist with the existing glow and outer shadow in native order glow, outer shadow, soft edge. Each field SHALL support creation, modification, deletion and recreation without losing the remaining fields or literal text. Removing all effects SHALL remove the direct effect list. No-op SHALL preserve original bytes; changed effects SHALL only modify the target ChartPart.

#### Scenario: Line and combo lifecycle
- **WHEN** a line or categorical combination chart edits/removes/restores soft edge and independently removes/restores its glow/shadow siblings
- **THEN** every fresh projection SHALL recover the requested combination and all non-target ZIP entries SHALL remain byte-identical.

#### Scenario: Authored theme-colored siblings
- **WHEN** an authored chart uses standard undeclared theme tokens for glow/shadow beside soft edge
- **THEN** resource-reference validation SHALL allow those direct effect colors while retaining rejection of wrong-kind declared grammar tokens and unsupported theme foreground colors.

### Requirement: Style precedence and vector propagation
Soft edge SHALL follow declared nested chart style precedence. Vector chart text SHALL inherit title defaults only when a run does not specify its own soft edge, independently of glow/shadow defaults.

#### Scenario: Explicit vector run zero
- **WHEN** a vector title has a nonzero default soft edge and one run explicitly sets zero
- **THEN** that run SHALL retain zero, other text SHALL use its appropriate defaults, and glow/shadow SHALL survive.

### Requirement: Unknown graphs preserve ownership
Missing or invalid native radius, duplicate/reordered effects, extra attributes/children and other unmodeled effects SHALL remain source-owned. This expands the preceding glow/shadow profile only by one supported final soft edge.

#### Scenario: Malformed imported effect
- **WHEN** a native rich chart run contains an unsupported effect graph
- **THEN** no-op SHALL retain the original bytes and an analytics edit SHALL fail without producing a file.
