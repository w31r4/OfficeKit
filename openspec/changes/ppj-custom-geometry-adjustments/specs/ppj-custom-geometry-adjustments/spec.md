## Purpose

Let PPJ retain and edit the native adjustment formulas that feed ordinary custom-geometry guides and text-layout references.

## ADDED Requirements

### Requirement: Custom adjustment formulas are distinct from preset values

Custom geometry SHALL accept optional geometry.adjustments as ordered name/formula pairs under the native bounded formula grammar. Preset geometry SHALL retain its existing integer-array contract. Custom adjustment order, formulas and dependent references SHALL survive authored export and fresh projection. Empty lists SHALL project as omission. Cross-list duplicate names, forward references and invalid arithmetic MUST reject; non-shape owners MUST reject this state.

#### Scenario: Adjustment feeds a guide
- **WHEN** a custom adjustment is referenced by a later adjustment and then an ordinary guide used by textRectangle
- **THEN** export and fresh projection SHALL retain the ordered lists and reference identities

### Requirement: Source custom adjustments support coherent list edits

Issued setGeometry authority SHALL permit geometry.adjustments add/change/removal on recognized custom shapes. Unchanged guides, paths, rectangle, text, frame and non-target package members SHALL be preserved. Source no-op SHALL preserve bytes. Removing a referenced adjustment MUST reject unless dependent state is coherently changed or removed in the same request.

#### Scenario: Value change and dependency cleanup
- **WHEN** separate original-source requests change an adjustment formula or remove it with its dependent graph
- **THEN** fresh projection SHALL retain the requested state and only the target SlidePart SHALL change

### Requirement: Preview preserves unverified formula limitations

Preview SHALL emit explicit limitations for adjustment-driven geometry it cannot fully render.

#### Scenario: Custom adjustment preview
- **WHEN** the preview does not implement the declared custom adjustment graph
- **THEN** the adjustment formula SHALL NOT be reported as fully supported
