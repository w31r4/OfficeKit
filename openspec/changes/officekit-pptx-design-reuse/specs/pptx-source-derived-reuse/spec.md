## ADDED Requirements

### Requirement: Reuse is ownership-checked
The source-derived reuse operation SHALL bind to a source SHA-256, selected
slide or component ID, expected revision, and a codec-proven ownership graph.

#### Scenario: Reuse a closed slide graph
- **WHEN** a selected imported slide has a complete uniquely owned descendant
  graph and the source hash matches
- **THEN** OfficeKit creates a distinct copy without changing the source slide
  or sharing mutable descendants

### Requirement: Ambiguous or opaque graphs fail closed
The reuse operation MUST reject shared descendants, unresolved relationships,
unsupported native topology, stale profiles, and any request that would expose
raw XML or arbitrary relationship edits.

#### Scenario: Template slide shares a media part
- **WHEN** the selected slide references media also owned by another slide
- **THEN** the reuse operation returns a structured unsupported-capability error
  before mutating the presentation

### Requirement: Reused output is reviewed independently
Every accepted reuse operation SHALL be followed by second import, package
footprint, structural, and render review before it is published or journaled.

#### Scenario: Reused slide is exported
- **WHEN** a source-derived copy is exported
- **THEN** OfficeKit proves the source is unchanged, the new graph is distinct,
  non-target parts remain unchanged, and review evidence is persisted
