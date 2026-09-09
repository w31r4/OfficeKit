## Purpose

Let PPJ custom shapes retain and edit named native geometry formulas, including dependencies used by their text-layout rectangle.

## ADDED Requirements

### Requirement: Ordered guide formulas are PPJ state

Custom shapes with literal paths SHALL accept optional geometry.guides as ordered name/formula pairs under the native bounded formula grammar. Guide names and formula meaning SHALL survive authored export and fresh projection; whitespace MAY normalize. Empty or omitted lists SHALL project without guides. Text rectangle references SHALL retain guide identity. Invalid names, duplicate or forward references and invalid evaluated graphs MUST reject. Masks and clips MUST reject this shape-only graph.

#### Scenario: Prior guide dependency
- **WHEN** a custom shape declares an inset formula followed by a formula depending on inset, and uses both in its text rectangle
- **THEN** export and fresh projection SHALL preserve guide order and rectangle references

### Requirement: Source guide lists support coherent edits

Issued setGeometry authority SHALL permit geometry.guides add/change/removal while preserving unchanged paths, rectangle references, text, frame and non-target package members. Source no-op SHALL preserve bytes. A removal that leaves a dangling reference MUST reject; coordinated removal of guides and their dependent rectangle SHALL be permitted.

#### Scenario: Formula change and safe removal
- **WHEN** separate original-source requests change a formula or remove the guide list with its dependent rectangle
- **THEN** fresh projection SHALL reflect the requested state and only the target SlidePart SHALL change

### Requirement: Unpainted guide effects remain visible as limitations

Preview SHALL retain machine-readable limitations for guide-driven geometry it cannot fully render.

#### Scenario: Formula-driven rectangle preview
- **WHEN** a preview cannot evaluate and lay out the declared guide-driven rectangle
- **THEN** the guide and rectangle state SHALL NOT be reported as fully supported
