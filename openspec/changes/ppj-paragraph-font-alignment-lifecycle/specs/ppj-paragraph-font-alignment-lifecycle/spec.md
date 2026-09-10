## Purpose

Express direct within-line font alignment in PPJ and edit it without changing
the surrounding paragraph, inline content or unrelated source package state.

## ADDED Requirements

### Requirement: Direct font alignment retains presence

PPJ SHALL accept `text.paragraphs[].style.fontAlignment` values `auto`, `top`,
`center`, `baseline` and `bottom`. Authoring SHALL use existing paragraph style
precedence and export/re-import SHALL preserve each value and absence separately.
Horizontal paragraph alignment, direction and text-box anchoring SHALL remain independent.

#### Scenario: Mixed font sizes with an explicit choice
- **WHEN** a paragraph declares a valid font alignment with differently sized runs
- **THEN** export and snapshot-free re-import retain that alignment and the run styles
- **AND** explicit `auto` remains distinct from no direct alignment

#### Scenario: Invalid input
- **WHEN** the alignment is an unknown string or a non-string value
- **THEN** compilation rejects it without an output artifact

### Requirement: Independent source lifecycle

Ordinary recognized text/shape source owners SHALL allow adding, setting,
removing and restoring the direct field under exact
`text.paragraphs[].style.fontAlignment` authority. Removing a style object that
contains only this field SHALL clear the direct attribute. Edits SHALL preserve
neighboring paragraphs, other native properties, inline content, relationships
and non-target package parts.

#### Scenario: Remove from original source and restore
- **WHEN** a fresh projection of original source bytes removes the field and a later edit restores it
- **THEN** re-import reflects absence and restoration and only the target paragraph attribute changes

#### Scenario: Missing field authority
- **WHEN** a request changes the field without its exact capability
- **THEN** the edit is rejected without an output artifact

### Requirement: Unknown native values stay source-owned

Unknown direct native font alignment SHALL remain omitted from semantic state,
preserved by no-op and unrelated edits, and reject replacement by a modeled value.

#### Scenario: Unrecognized native token
- **WHEN** a paragraph contains an unknown alignment token
- **THEN** a no-op preserves the original file and an independent default-bold edit preserves that token
- **AND** assigning a modeled font alignment is rejected

### Requirement: Discoverability and preview limits

The schema, Help, registry and generated/manual guidance SHALL describe the same
field lifecycle. Preview SHALL emit an explicit partial/unmapped diagnostic for
direct font alignment until font metrics and line layout implement it.

#### Scenario: Explicit automatic font alignment
- **WHEN** preview assesses `fontAlignment: "auto"` or another declared value
- **THEN** it reports the unsupported alignment and does not silently treat it as painted
