# PPJ missing chart points

## ADDED Requirements

### Requirement: Authored PPJ preserves missing chart observations

OfficeKit SHALL compile a bounded `null` Y value as a missing observation
without coercing it to zero or changing the logical point count.

#### Scenario: Build and reproject a line gap

- **GIVEN** a valid PPJ line chart with one `null` series value
- **WHEN** OfficeKit builds and reprojects the presentation
- **THEN** the native cache declares the complete point count
- **AND** the missing index has no native numeric point
- **AND** the projected PPJ contains `null` at the same index

### Requirement: Missing caches remain bounded and deterministic

OfficeKit SHALL accept only unique, ordered, in-range missing Y indexes with
finite present values.

#### Scenario: Reject an irregular sparse cache

- **GIVEN** a native chart cache with duplicate, unordered, out-of-range, or
  non-finite present points
- **WHEN** OfficeKit imports the presentation
- **THEN** the chart is not exposed as a typed editable chart
- **AND** the source graph remains preserved

### Requirement: Source-bound edits preserve missing topology

OfficeKit SHALL NOT add or remove missing points through the bounded
source-bound chart data capability.

#### Scenario: Edit a present value beside a gap

- **GIVEN** a typed imported chart with a fixed missing-point set
- **WHEN** an issued chart-data edit changes one present numeric value
- **THEN** the edit succeeds without changing the missing-point set
- **AND** an attempted present-to-missing or missing-to-present mutation is
  rejected before package mutation
