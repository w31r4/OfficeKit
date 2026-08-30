## ADDED Requirements

### Requirement: PPJ SHALL express one bounded treemap forest

The language SHALL accept `chartType: "treemap"` with unique string categories,
positive values and aligned nullable `parents`. It SHALL reject missing
parents, cycles, excessive depth, excessive roots/nodes and inconsistent parent
totals before native output.

#### Scenario: Valid hierarchy

- **WHEN** every non-leaf value equals its direct-child sum and the parent graph is a bounded forest
- **THEN** validation succeeds deterministically

#### Scenario: Cyclic hierarchy

- **WHEN** any parent chain returns to a visited node
- **THEN** validation rejects the PPJ without output

### Requirement: NativeAOT SHALL compile an editable vector treemap

The compiler SHALL emit one deterministic native group containing editable
node rectangles and bounded labels. It SHALL NOT add a raster image or claim a
ChartPart.

#### Scenario: Multi-root build

- **WHEN** the input contains multiple valid roots
- **THEN** root areas follow their declared values and descendants remain inside their parent regions

### Requirement: Recovery SHALL distinguish semantic and native state

An authored PPTX with embedded PPJ SHALL restore the exact treemap program. If
the embedded program is removed, projection SHALL expose an ordinary editable
group and SHALL NOT infer treemap semantics.

#### Scenario: Snapshot removed

- **WHEN** a compiled treemap PPTX no longer contains the OfficeKit program
- **THEN** import returns the native group and no treemap chart node

### Requirement: Agent guidance SHALL expose purpose and limits

The generated PPJ manual and focused chart guide SHALL document the parent
channel, total invariant, style fields, editability and non-ChartPart animation
boundary.

#### Scenario: Fresh Agent searches for hierarchical part-to-whole

- **WHEN** an Agent reads the chart reference
- **THEN** it can find a minimal treemap example and its finite restrictions
