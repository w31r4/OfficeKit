## Purpose

Make PPJ direct text column count editable and removable without losing surrounding text, spacing, direction or source package state.

## ADDED Requirements

### Requirement: Direct column count presence lifecycle

The system SHALL distinguish absent columns from explicit integers 1 through 16 for supported source-bound text, shape text, master/layout placeholders and table cells. Deletion SHALL remove native numCol, preserve other properties and reproject as absent; restoration SHALL retain explicit values.

#### Scenario: Delete and restore column count
- **WHEN** a projected count is removed and subsequently restored as 1, 3 or 16
- **THEN** native attributes and fresh projection reflect each operation, gap and direction remain unchanged, no-op bytes remain identical and non-target package entries are preserved

### Requirement: Guarded simple style removal

The system SHALL permit whole-style removal when it contains only the supported removable body properties including columns, while retaining authority checks and rejection for other fields.

#### Scenario: Remove a count-only or combined style
- **WHEN** a shape or table cell loses its whole removable style
- **THEN** its direct properties disappear and explicit single-column text can be restored, including after table text compaction

### Requirement: Honest preview boundary

The preview SHALL retain an explicit unsupported layout diagnostic for columns, including the value 1.

#### Scenario: Explicit single-column preview input
- **WHEN** input assessment encounters columns 1
- **THEN** it reports the field as not fully supported by preview layout
