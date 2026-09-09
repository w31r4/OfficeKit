## Purpose

Complete PPJ text-body wrapping presence semantics so source edits distinguish explicit wrapping choices from inherited/default behavior without changing unrelated content.

## ADDED Requirements

### Requirement: Wrap presence lifecycle
The system SHALL preserve square and none as explicit wrap values. Removing a previously projected wrap field SHALL remove native bodyPr wrap under existing source edit authority for supported text, shape, master/layout placeholder and table-cell owners.

#### Scenario: Delete then restore wrapping
- **WHEN** the caller deletes wrap and subsequently restores none or square
- **THEN** native wrap and fresh projection SHALL respectively be absent or contain the explicit requested value.

### Requirement: Simple style removal and preservation
A style containing only wrap, upright, rotation, columnDirection and/or verticalText SHALL be removable. Existing capability and other-field deletion guards SHALL remain enforced. No-op SHALL preserve source bytes; isolated edits SHALL preserve surrounding XML, text/frame state and non-target ZIP entries.

#### Scenario: Restore compact table text
- **WHEN** deleting a simple style compacts a table cell to plain text
- **THEN** equivalent structured text with unchanged native topology SHALL restore explicit wrap=none.

### Requirement: Honest preview coverage
Preview SHALL retain an explicit field-level layout diagnostic for wrap choices it cannot fully render.

#### Scenario: Assess explicit no wrapping
- **WHEN** the input text style specifies wrap=none
- **THEN** preview assessment SHALL report the layout limitation.
