## Purpose

This capability gives fixed-topology PPJ table cells a source-preserving way
to change a static DrawingML field identity while retaining the cell's cached
text and native table structure.

## ADDED Requirements

### Requirement: project static table-cell field identities

The PPJ projector MUST expose
`table.rows[].cells[].text.paragraphs[].runs[].field.id` through a
`setTextField` capability and a `tableTextFieldId` native leaf when an
editable rectangular table cell contains a valid non-automatic DrawingML
field with a brace-wrapped UUID identity.

#### Scenario: table field identity receives a bounded native leaf

- **WHEN** an imported editable table cell contains a valid static `a:fld`
  with an ID, type, and cached text
- **THEN** the cell's native reference contains the exact table-cell field
  path and a leaf whose value is the source field ID

### Requirement: edit only the table field identity token

The source-bound compiler MUST accept a changed valid table-cell field ID only
when the fresh source authority, cell index, field index, expected ID, field
type, cached text, and automatic state match, and MUST change only the owning
slide's direct `a:fld/@id` attribute.

#### Scenario: static table field identity round-trips

- **WHEN** a projected table field ID changes to another valid UUID while its
  type, cached text, automatic state, inline topology, table grid, and merges
  stay unchanged
- **THEN** compilation succeeds, changes only the owning SlidePart, and a
  fresh projection reports the new ID with the same type and cached text

### Requirement: reject unsafe table-field identity changes

The compiler MUST fail closed for automatic or invalid IDs, stale leaf
authority, cached-text or type changes in the same edit, automatic-state or
field-topology changes, and any non-rectangular or unsupported table graph.

#### Scenario: combined table field identity and cached text change is rejected

- **WHEN** a table field ID and its cached display text are changed in one
  source-bound edit
- **THEN** source-bound compilation fails and emits no output mutation
