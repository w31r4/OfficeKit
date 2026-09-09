## Purpose

Complete vertical text overflow presence and restoration for source-bound PPJ while preserving horizontal overflow and unrelated presentation content.

## ADDED Requirements

### Requirement: Vertical overflow lifecycle
The system SHALL preserve explicit overflow, ellipsis and clip. Removing a projected verticalOverflow SHALL remove bodyPr vertOverflow under existing authority for supported text, shape, master/layout placeholder and table-cell owners.

#### Scenario: Delete and restore each value
- **WHEN** verticalOverflow is removed then restored to each supported value
- **THEN** native attributes and fresh projection SHALL respectively be absent or retain the explicit value, with horizontalOverflow unchanged.

### Requirement: Style removal and preservation
A style containing only verticalOverflow, horizontalOverflow, upright, rotation, columnDirection, verticalText and/or wrap SHALL be removable. Other-field deletion and missing authority SHALL remain rejected. No-op SHALL preserve source bytes; isolated edits SHALL preserve surrounding XML, text/frame state and non-target ZIP entries.

#### Scenario: Restore compact table text
- **WHEN** deletion of a simple style compacts a table cell to plain text
- **THEN** equivalent structured text with unchanged native topology SHALL restore explicit verticalOverflow.

### Requirement: Honest preview limitation
Preview SHALL retain a field-level diagnostic for unsupported verticalOverflow rendering.

#### Scenario: Preview ellipsis
- **WHEN** a text style specifies verticalOverflow=ellipsis
- **THEN** assessment SHALL report the unsupported layout state.
