# PPJ custom-geometry arcs

## ADDED Requirements

### Requirement: PPJ authors native editable arcs

PPJ SHALL express a bounded elliptical arc as a typed custom-geometry command
using view-box radii and degree angles, and SHALL compile it into the native
DrawingML arc primitive.

#### Scenario: Literal authored arc

- **WHEN** an `arcTo` follows an established current point with positive radii
  and a non-zero sweep of at most one turn
- **THEN** the compiler SHALL write an editable `a:arcTo`
- **AND** import without an embedded program SHALL recover the typed command.

#### Scenario: Arc has no current point

- **WHEN** an arc is the first path command or otherwise has no current point
- **THEN** PPJ validation SHALL reject the program before native output changes.

#### Scenario: Native arc depends on formulas

- **WHEN** an imported custom arc uses native guide or adjustment references
- **THEN** PPJ SHALL preserve the source geometry as opaque
- **AND** SHALL NOT guess literal radii or angles.

