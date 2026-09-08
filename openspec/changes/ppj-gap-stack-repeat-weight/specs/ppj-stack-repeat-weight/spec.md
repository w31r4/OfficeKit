## Purpose

Provides a bounded, explicit weight owner for authored component repeats so
unequal stack slots can be represented and projected without introducing a
general layout solver.

## ADDED Requirements

### Requirement: Weighted repeat slots are explicit and deterministic

An authored component repeat MAY provide `layout.weights` together with a
horizontal or vertical layout direction. The weights MUST contain exactly one
positive finite number per repeat item. The available main-axis extent MUST be
the container extent after subtracting the configured gap between adjacent
items, and each item slot MUST receive its share of that extent proportional to
its weight. The configured gap MUST remain between adjacent slots.

#### Scenario: Horizontal weighted repeat

- **WHEN** a horizontal repeat has three items, gap `12`, and weights `[1,2,1]`
- **THEN** the three projected slots have widths in the ratio `1:2:1`, remain
  inside the instance frame, and have exactly `12` units between neighbors

#### Scenario: Vertical weighted repeat

- **WHEN** a vertical repeat has weighted items and a nonzero gap
- **THEN** the item heights follow the declared weight ratio and the gap remains
  between consecutive items on the vertical axis

### Requirement: Invalid weighted layout fails closed

The program MUST reject weights used with `grid` or `flow`, weights combined
with `anchor`, a weight count different from the repeat item count, or a weight
that is non-positive, non-finite, or outside the bounded numeric range. The
diagnostic MUST identify the invalid `weights` declaration or offending weight
entry.

#### Scenario: Weight count does not match items

- **WHEN** a repeat contains three items and two weights
- **THEN** validation fails with a diagnostic at the repeat layout weights

#### Scenario: Zero or negative weight is supplied

- **WHEN** a weighted repeat contains a zero or negative weight
- **THEN** validation fails with a diagnostic at that weight entry and no
  weighted frames are authored

#### Scenario: Weight is used with an incompatible layout

- **WHEN** weights are supplied for a grid or flow repeat, or alongside an
  anchor
- **THEN** validation fails rather than dropping the unsupported declaration

### Requirement: Existing equal repeats remain compatible

When `layout.weights` is omitted, horizontal and vertical repeats MUST retain
their existing equal-slot behavior. Weighted authored expansion MUST lower to
ordinary PPJ elements and frames; it MUST NOT claim source-bound layout or
solver ownership.

#### Scenario: Existing repeat omits weights

- **WHEN** an existing horizontal, vertical, grid, or flow repeat has no
  weights field
- **THEN** it validates and projects with its prior equal/grid/flow behavior

#### Scenario: Weighted output is reprojected

- **WHEN** a valid weighted authored program is compiled to PPTX and projected
  after embedded PPJ metadata is removed
- **THEN** the result contains ordinary projected element frames with the
  weighted geometry and no new source-bound layout capability
