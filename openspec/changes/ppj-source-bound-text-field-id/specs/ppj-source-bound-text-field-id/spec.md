## Purpose

This capability gives PPJ a source-preserving identity field for direct static
DrawingML text fields while keeping host-evaluated and composite field
semantics outside the editable contract.

## ADDED Requirements

### Requirement: project direct static field identity

The projector MUST expose `text.paragraphs[].runs[].field.id` as a
`textFieldId` native leaf for each valid non-automatic direct DrawingML field
in an editable text body. The leaf value MUST be the exact existing
brace-wrapped UUID and its native index MUST identify the field ordinal in the
owning text shape.

#### Scenario: imported static field receives an identity leaf

- **GIVEN** an imported shape contains a direct non-automatic `a:fld` with a
  valid brace-wrapped UUID, a type, and cached text
- **WHEN** the shape is projected to PPJ
- **THEN** the native reference contains `textFieldId` at the field path with
  the exact source ID

### Requirement: compile one direct field identity token

The source-bound compiler MUST accept a changed valid brace-wrapped UUID only
when the fresh native reference, source hash, field ordinal, and expected ID
match. It MUST replace only the owning slide's direct `a:fld/@id` token.

#### Scenario: identity change preserves field content and package scope

- **GIVEN** a projected static field whose type, cached text, automatic state,
  and inline topology are unchanged
- **WHEN** its `textFieldId` changes to another valid UUID
- **THEN** compilation succeeds, changes only the owning slide part, and a
  fresh projection reports the new ID with the same type and cached text

### Requirement: reject stale, invalid, and unsupported identity edits

The compiler MUST fail closed for a malformed or automatic field ID, a stale
source or expected value, a field outside the issued direct-field profile, or
any edit that changes field type, cached text, automatic state, or inline
topology alongside the identity edit.

#### Scenario: stale or malformed identity is rejected

- **GIVEN** a `textFieldId` leaf whose source hash or expected ID no longer
  matches, or whose requested value is not a brace-wrapped UUID
- **WHEN** a source-bound edit is compiled
- **THEN** compilation fails before changing source bytes

#### Scenario: automatic and table identities remain source-owned

- **GIVEN** an automatic DrawingML field or a field nested in a table
- **WHEN** it is projected to PPJ
- **THEN** no editable `textFieldId` leaf is issued for that field
