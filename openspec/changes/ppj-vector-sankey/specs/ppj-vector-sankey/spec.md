## ADDED Requirements

### Requirement: PPJ SHALL express one bounded Sankey DAG

The language SHALL accept `chartType: "sankey"` with a stable node catalog and
aligned positive source, target and flow channels. It SHALL reject unknown
endpoints, self edges, duplicate edges, cycles, disconnected nodes, excessive
budgets and non-conserving internal nodes before native output.

#### Scenario: Valid conserved flow

- **WHEN** a finite directed graph is acyclic and every internal node conserves flow
- **THEN** validation succeeds deterministically

### Requirement: NativeAOT SHALL compile editable nodes and ribbons

The compiler SHALL assign topological columns, stack each edge consistently at
both endpoints and emit editable rectangles, custom-geometry ribbons and text.
It SHALL NOT introduce a PNG or claim a ChartPart.

#### Scenario: Native ribbon

- **WHEN** one valid flow connects adjacent or non-adjacent columns
- **THEN** the closed ribbon uses bounded cubic paths behind its endpoint nodes

### Requirement: Recovery SHALL not infer graphs from shapes

Embedded PPJ SHALL restore exact Sankey intent. Without it, import SHALL expose
an ordinary editable group and SHALL NOT infer graph semantics.

#### Scenario: Snapshot removed

- **WHEN** the private program is absent
- **THEN** import returns custom shapes and no Sankey chart node

### Requirement: Agent guidance SHALL expose relationship purpose and limits

The generated PPJ manual and focused chart guide SHALL document node/edge
channels, conservation, native editability and animation limits.

#### Scenario: Fresh Agent searches for directed flow

- **WHEN** an Agent reads the chart reference
- **THEN** it can distinguish Sankey from a process diagram and find one example
