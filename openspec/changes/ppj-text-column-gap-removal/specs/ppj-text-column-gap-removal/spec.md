## Purpose

Complete columnGap deletion and numeric restoration in source-bound PPJ while preserving column count, direction and unrelated presentation content.

## ADDED Requirements

### Requirement: Column gap presence lifecycle
The system SHALL preserve columnGap numbers from 0 through 10000 points, including explicit zero. Removing a projected field SHALL remove bodyPr spcCol under existing authority for supported text, shape, master/layout placeholder and table-cell owners.

#### Scenario: Remove and restore numeric spacing
- **WHEN** columnGap is removed then restored to zero, a fractional value or the maximum
- **THEN** native spacing and fresh projection SHALL respectively be absent or retain the requested numeric value with existing unit conversion, while columns and columnDirection remain unchanged.

### Requirement: Simple style removal and preservation
A style containing only columnGap, verticalAlignment, upright, rotation, columnDirection, verticalText, wrap, horizontalOverflow and/or verticalOverflow SHALL be removable. Other-field deletion and missing authority SHALL remain rejected. No-op SHALL preserve source bytes; isolated edits SHALL preserve surrounding XML, text/frame state and non-target ZIP entries.

#### Scenario: Restore compact table text
- **WHEN** deleting a simple style compacts table text
- **THEN** equivalent structured text with unchanged native topology SHALL restore explicit columnGap=0.

### Requirement: Honest preview limitation
Preview SHALL retain an explicit field-level limitation for unsupported column layout.

#### Scenario: Assess explicit zero spacing
- **WHEN** text style specifies columnGap=0
- **THEN** preview assessment SHALL report its layout limitation.
