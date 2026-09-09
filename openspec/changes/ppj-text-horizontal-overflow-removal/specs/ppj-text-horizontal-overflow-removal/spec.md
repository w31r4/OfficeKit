## Purpose

Complete horizontal text overflow presence and restoration in source-bound PPJ while retaining independent vertical overflow and unrelated presentation content.

## ADDED Requirements

### Requirement: Horizontal overflow presence lifecycle
The system SHALL preserve explicit overflow and clip. Removing a projected horizontalOverflow SHALL remove native bodyPr horzOverflow under existing authority for supported text, shape, master/layout placeholder and table-cell owners.

#### Scenario: Delete and restore overflow
- **WHEN** the caller removes horizontalOverflow then restores overflow or clip
- **THEN** native attributes and fresh projection SHALL respectively be absent or contain the explicit value, with verticalOverflow unchanged.

### Requirement: Style removal and preservation
A style containing only horizontalOverflow, upright, rotation, columnDirection, verticalText and/or wrap SHALL be removable. Other-field deletion and missing authority SHALL remain rejected. No-op SHALL preserve source bytes and isolated edits SHALL preserve surrounding XML, text/frame state and non-target ZIP entries.

#### Scenario: Restore compact table text
- **WHEN** removal of a simple style compacts table text
- **THEN** equivalent structured text with unchanged native topology SHALL restore explicit horizontalOverflow.

### Requirement: Honest preview limitation
Preview SHALL retain a field-level layout diagnostic for unsupported horizontalOverflow rendering.

#### Scenario: Preview explicit clipping
- **WHEN** a text style specifies horizontalOverflow=clip
- **THEN** assessment SHALL report the unsupported layout state.
