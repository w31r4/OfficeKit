## ADDED Requirements

### Requirement: PPJ SHALL express one bounded sunburst forest

The language SHALL accept `chartType: "sunburst"` with unique string categories,
positive values and aligned nullable parents. It SHALL reject missing parents,
cycles, inconsistent totals, more than 96 nodes, more than 16 roots or hierarchy
depth above six before native output.

#### Scenario: Valid radial hierarchy

- **WHEN** the finite forest and direct-child totals are valid
- **THEN** validation succeeds deterministically

### Requirement: NativeAOT SHALL compile editable radial sectors

The compiler SHALL partition angle by value and depth by ring, then emit one
native group of editable custom-geometry sectors and bounded text. It SHALL NOT
add a raster image or claim a ChartPart.

#### Scenario: Sector path

- **WHEN** a node owns a nonzero angular interval
- **THEN** its sector uses bounded cubic paths and stays inside its declared ring

### Requirement: Recovery SHALL not infer semantics from arcs

Embedded PPJ SHALL restore exact sunburst intent. Without it, import SHALL
expose an ordinary editable group and SHALL NOT infer the parent/value graph.

#### Scenario: Snapshot removed

- **WHEN** the private program is absent
- **THEN** import returns custom shapes and no sunburst chart node

### Requirement: Agent guidance SHALL expose radial purpose and limits

The generated PPJ manual and focused chart guide SHALL document the shared
hierarchy channel, radial style, editability and animation boundary.

#### Scenario: Fresh Agent searches for radial hierarchy

- **WHEN** an Agent reads the chart reference
- **THEN** it can distinguish sunburst from treemap and find a minimal example
