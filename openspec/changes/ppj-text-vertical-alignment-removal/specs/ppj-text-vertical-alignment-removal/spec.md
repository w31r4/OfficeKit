## Purpose

Complete direct verticalAlignment deletion and restoration in source-bound PPJ without losing anchorCenter or unrelated presentation content.

## ADDED Requirements

### Requirement: Vertical alignment presence lifecycle
The system SHALL preserve explicit top/middle/bottom. Removing projected verticalAlignment SHALL remove native bodyPr anchor under existing authority for supported text, shape, master/layout placeholder and table-cell owners.

#### Scenario: Remove and restore every alignment
- **WHEN** verticalAlignment is removed then restored to each supported value
- **THEN** native anchor and fresh projection SHALL respectively be absent or retain the explicit value, mapping middle to native center and preserving anchorCenter.

### Requirement: Simple style removal and preservation
A style containing only verticalAlignment, upright, rotation, columnDirection, verticalText, wrap, horizontalOverflow and/or verticalOverflow SHALL be removable. Other-field deletion and missing authority SHALL remain rejected. No-op SHALL preserve bytes; isolated edits SHALL preserve surrounding XML, text/frame state and non-target ZIP entries.

#### Scenario: Restore compact table text
- **WHEN** deleting a simple style compacts a table cell to plain text
- **THEN** equivalent structured text with unchanged native topology SHALL restore explicit verticalAlignment=top.

### Requirement: Honest preview limitation
Preview SHALL retain explicit layout limitations for verticalAlignment that it cannot fully render.

#### Scenario: Assess middle alignment
- **WHEN** a text style specifies verticalAlignment=middle
- **THEN** preview assessment SHALL report its unsupported layout state.
