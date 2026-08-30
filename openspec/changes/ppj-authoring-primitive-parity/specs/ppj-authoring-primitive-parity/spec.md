## ADDED Requirements

### Requirement: Complete finite preset vocabulary
PPJ SHALL accept and round-trip `upArrow` and `lineInv` through the same shared
preset registry used by authored and imported Presentation shapes.

#### Scenario: Agent authors either missing preset
- **WHEN** a shape declares one of the two finite preset names
- **THEN** build writes the matching DrawingML preset and second import returns
  the same PPJ geometry

### Requirement: Bounded table style inheritance
PPJ SHALL expand finite table-level cell styles deterministically before native
table compilation.

#### Scenario: Row and column roles overlap
- **WHEN** a cell is covered by body, row and column styles
- **THEN** property-level merging follows the declared `rowOverColumn` order
  and an explicit cell property wins last

### Requirement: Complete bounded Sankey positioning
PPJ SHALL support left, right and justified node alignment plus declared-node
color overrides in the existing vector Sankey profile.

#### Scenario: Agent selects right alignment
- **WHEN** the declared graph is a valid flow-conserving DAG
- **THEN** node columns are derived from reverse topological depth and the
  output remains deterministic and editable

#### Scenario: Color override names an unknown node
- **WHEN** a node-color key does not match a declared category
- **THEN** validation rejects before native writing
