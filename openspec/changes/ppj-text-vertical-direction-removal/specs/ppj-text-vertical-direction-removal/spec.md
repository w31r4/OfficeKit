## Purpose

Complete verticalText source-bound presence and restoration for existing PPJ text owners without losing native text direction or unrelated presentation content.

## ADDED Requirements

### Requirement: Text direction presence lifecycle
The system SHALL preserve horizontal, vertical and vertical270 as explicit direct values, and remove native bodyPr vert when a projected verticalText field is removed under existing edit authority. This SHALL apply to supported text, shape, master/layout placeholder and table-cell owners.

#### Scenario: Remove and restore text direction
- **WHEN** the caller removes a projected verticalText then restores each supported value
- **THEN** native vert and fresh projection SHALL respectively be absent or contain the matching explicit direction, including horizontal.

### Requirement: Simple style removal and preservation
A style containing only verticalText, upright, rotation and/or columnDirection SHALL be removable. Other whole-style deletion and missing authority SHALL remain rejected. Source no-op SHALL preserve bytes; direction edits SHALL preserve other body properties, frame, text topology and non-target ZIP entries.

#### Scenario: Restore compact table text
- **WHEN** removal of a simple style compacts a table cell to plain text
- **THEN** equivalent structured text with unchanged native topology SHALL restore an explicit verticalText value.

### Requirement: Explicit preview limitation
Preview SHALL report its unsupported text-layout state for verticalText instead of silently treating the field as fully rendered.

#### Scenario: Preview explicit horizontal direction
- **WHEN** a text style explicitly sets horizontal verticalText
- **THEN** input assessment SHALL retain the field-level layout diagnostic.
