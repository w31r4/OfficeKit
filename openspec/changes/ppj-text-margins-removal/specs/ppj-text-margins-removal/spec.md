## Purpose

Allow PPJ users to remove and restore direct text margins while preserving other inset edges, surrounding text and source ownership.

## ADDED Requirements

### Requirement: Independent inset presence

The system SHALL preserve explicit left/top/right/bottom margins from 0 through 10000 points and distinguish them from absence on supported source-bound text, shape, master/layout placeholder and table-cell owners.

#### Scenario: Delete and restore one edge
- **WHEN** one projected edge is deleted and restored as zero or a fractional value
- **THEN** its native inset attribute disappears and returns with the correct value, fresh projection agrees, other edges and package state are preserved, and unchanged input is byte-identical

### Requirement: Nested and whole style deletion

The system SHALL remove all previously projected direct margins when the margins object is removed or emptied. A style composed only of supported removable properties including margins SHALL support whole removal and restoration. Other-property and edit-authority guards SHALL remain enforced.

#### Scenario: Delete a margins object or a simple style
- **WHEN** margins are emptied, removed, or deleted with a simple style
- **THEN** all four native direct insets are absent and explicit zero can be restored, including compact table text

### Requirement: Preview limitations stay visible

Preview assessment SHALL retain field-addressable layout limitations for explicit margins, including zero.

#### Scenario: Zero inset preview input
- **WHEN** a zero margin is assessed
- **THEN** unsupported layout is not declared fully supported
