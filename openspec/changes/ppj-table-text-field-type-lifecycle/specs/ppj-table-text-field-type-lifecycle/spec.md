## Purpose

This capability gives fixed-topology PPJ table cells a source-preserving way to
change a static DrawingML field type while keeping the cell's existing text,
identity, and table structure intact.

## ADDED Requirements

### Requirement: project static table-cell field types

The PPJ projector MUST expose
`table.rows[].cells[].text.paragraphs[].runs[].field.type` through a
`setTextField` capability and a native leaf when a rectangular editable table
cell contains a valid non-automatic DrawingML field.

#### Scenario: table field receives a bounded native leaf

- **WHEN** an imported editable table cell contains a valid static `a:fld`
  with an ID and cached text
- **THEN** the cell's table element reference contains the exact table-cell
  field path and a leaf whose value is the source field type

### Requirement: edit only the table field type token

The source-bound compiler MUST accept a changed valid non-automatic table-cell
field type only when the fresh source authority, cell index, field index, and
expected token match, and MUST change only the owning slide's direct
`a:fld/@type` attribute.

#### Scenario: static table field type round-trips

- **WHEN** a projected table field changes from one static type token to
  another while its ID, cached text, automatic state, inline topology, table
  grid, and merges stay unchanged
- **THEN** compilation succeeds, changes only the owning SlidePart, and a
  fresh projection reports the new type with the same ID and cached text

### Requirement: reject unsafe table-field changes

The compiler MUST fail closed for automatic or invalid type tokens, stale leaf
authority, field identity changes, cached-text changes in the same edit,
automatic-state changes, paragraph/run or field topology changes, and any
non-rectangular or unsupported table cell.

#### Scenario: automatic table field type change is rejected

- **WHEN** a table field's type is changed to a bounded automatic field type
- **THEN** source-bound compilation fails and emits no output mutation
