## Purpose

Make full-span chart character reflection independently expressible and editable without losing optional native geometry or opacity state.

## ADDED Requirements

### Requirement: Typed reflection presence
Chart text styles SHALL accept reflection with optional blur (0..1000 points), distance (0..100000 points), angle (-360..360 degrees), startOpacity and endOpacity (0..1 or opacity grammar token). Empty object SHALL retain a full-span reflection; omission SHALL remove it. Native optional scalar presence and explicit zero SHALL survive wire transport, writing and projection. Named style field precedence and vector explicit run overrides SHALL apply.

#### Scenario: Empty and explicit reflection
- **WHEN** a chart character style declares an empty reflection or a reflection with explicit zero scalars
- **THEN** native output and fresh projection retain the distinction, with start/end positions fixed at 0/100000

### Requirement: Independent source-bound reflection lifecycle
The system SHALL create, modify, delete and recreate reflection independently beside glow, inner shadow, outer shadow and soft edge, in that order. Source-bound no-op SHALL retain exact source bytes; reflection edits SHALL preserve literal text and every non-target ZIP part.

#### Scenario: Line and combo chart character owners
- **WHEN** fresh native projections of line and combo charts edit reflection in title, legend, axes, data labels and rich trendline styles
- **THEN** the candidate changes only the chart part, preserves siblings, and reprojects the selected value and presence

### Requirement: Unsupported reflection topology remains source-owned
Non-full-span, transformed, duplicate, reordered, out-of-bound or extended native reflection owners SHALL remain source-owned. Their no-op SHALL preserve original bytes and unsupported typed edits SHALL fail closed.

#### Scenario: Transformed or partial-span native reflection
- **WHEN** a source reflection has a transform attribute, noncanonical span or unknown child
- **THEN** it gains no chart analytics edit capability and an attempted analytics edit is rejected

### Requirement: Ordinary reflection angle leaf uses degrees
The existing textReflectionDirectionDegrees leaf SHALL convert degree values, including fractional values, to the declared native angle scale when proving the source and applying edits.

#### Scenario: Fractional source angle and edit
- **WHEN** a projected ordinary text reflection angle changes from 45.5 degrees to 90.25 degrees
- **THEN** the source proof accepts the unchanged native baseline and writes 5415000 native angle units, preserving non-target package entries
