## Context

The table profile already preserves rectangular cell topology and exposes a
`tableTextFieldType` leaf whose `TextLeafIndex` selects a flattened physical
cell and whose `NativeLeafIndex` selects a field in source order. Ordinary
shape text uses the same `setTextField` operation for a static field ID, but
table field IDs are currently rejected as identity changes.

## Goals / Non-Goals

**Goals:**

- issue one identity leaf only for an existing valid non-automatic table field;
- bind the cell and field ordinals to the existing source hash and capability;
- splice one direct `a:fld/@id` attribute and prove a second projection.

**Non-Goals:**

- field type, cached display text, automatic evaluation, relationships, or
  inline/table topology changes;
- adding IDs to automatic, malformed, irregular, or unsupported field graphs;
- changing the protobuf contract or implementing host refresh behavior.

## Decisions

Use a distinct `tableTextFieldId` native leaf while reusing `setTextField`.
`TextLeafIndex` identifies the flattened physical cell and
`NativeLeafIndex` identifies the field ordinal, including automatic fields in
the count. The projector emits the ID leaf before the existing type leaf for a
valid static field, so both leaves address the same source field without
changing the table text path.

The table mutation pass ignores field ID and type when deciding whether a cell
text body changed, then emits separate native mutations for changed ID and
type tokens. A field-ID mutation is accepted only when both sides retain valid
UUIDs, a valid non-automatic type, the same cached text and automatic state,
and the exact capability. This allows independent ID edits while keeping
combined text/identity edits fail-closed.

The edit-plan proof reads the same table cell and field ordinal from the
reopened source tree and replaces only the quoted `id` attribute in the field
start tag. It reuses the existing XML token and changed-part checks rather than
serializing a table or text body.

## Risks / Trade-offs

- [Field ordinal drift] → Count every source field, including automatic ones,
  and retain both cell and field indices in the native leaf binding.
- [Accidental table flattening] → Require the existing rectangular editable
  table and bounded mixed-run text body before issuing the leaf.
- [False semantic parity] → Keep automatic refresh, field relationships, and
  host PowerPoint behavior outside the profile and document the source-bound
  boundary.
