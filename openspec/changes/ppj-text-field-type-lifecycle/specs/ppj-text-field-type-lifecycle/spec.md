# PPJ static text-field type lifecycle

## ADDED Requirements

### Requirement: project editable static field types
The projector MUST expose `text.paragraphs[].runs[].field.type` through a
`setTextField` capability and a `textFieldType` native leaf when the source
field is a valid non-automatic DrawingML field in an editable text body.

#### Scenario: imported static field receives a bounded leaf
- **GIVEN** an imported shape contains a valid static `a:fld` with an ID and
  cached `a:t` text
- **WHEN** the shape is projected to PPJ
- **THEN** its native reference contains the exact field path and a leaf whose
  value is the field type

### Requirement: compile one static field type token
The source-bound compiler MUST accept a changed static field type only when
the fresh native reference, source hash, field ordinal, and expected type
match, and MUST patch only the owning slide's direct `a:fld/@type` token.

#### Scenario: type change preserves field identity and package scope
- **GIVEN** a projected static field with unchanged ID, automatic marker,
  cached text, and inline topology
- **WHEN** its type changes to another valid static token
- **THEN** compilation succeeds, changes only the owning slide part, and a
  fresh projection reports the new type with the same ID and cached text

### Requirement: reject automatic or topology changes
The compiler MUST fail closed when the requested field type is automatic, when
the source or requested field changes ID, cached text, automatic state, run
topology, or field topology, or when the field falls outside the issued leaf
profile.

#### Scenario: automatic field type change is rejected
- **GIVEN** a field whose type is one of the bounded automatic field types
- **WHEN** its PPJ type is changed
- **THEN** source-bound compilation fails without changing the source bytes
