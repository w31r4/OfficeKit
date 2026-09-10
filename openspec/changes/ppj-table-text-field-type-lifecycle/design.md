## Context

The table profile already preserves fixed paragraph/run topology and exposes
field runs for cached-text edits. Its native reference currently has one
`tableCellText` leaf per physical cell; ordinary shape text has a separate
static field-type leaf and token splice. See the proposal and delta spec for
the externally visible contract.

## Goals / Non-Goals

**Goals:**

- Reuse the existing source-bound leaf proof for one physical table cell and
  one field ordinal.
- Keep the public table text path explicit and leave all other cell state
  source-preserved.
- Make a one-attribute SlidePart edit and prove a second projection.

**Non-Goals:**

- Changing table grid/merge topology, cached display text, field IDs, or
  automatic evaluation/refresh.
- Adding fields to empty cells, irregular tables, external field graphs, or
  fields hidden by an unsupported text-body profile.

## Decisions

Use a distinct `tableTextFieldType` native leaf kind while reusing the
`setTextField` operation. The leaf's `TextLeafIndex` identifies the flattened
physical cell and its `NativeLeafIndex` identifies the field ordinal within
that cell; automatic fields still advance the ordinal but receive no leaf.
This avoids overloading the ordinary shape field's coordinate interpretation
and keeps stale binding errors explicit.

The table projector issues leaves only from a preserved structured text body
whose field type passes the existing printable-token grammar and is not a
recognized automatic type. The compiler's existing native-reference diff
collector turns a changed leaf into a mutation, while table text comparison
ignores a field type only when every other text property is unchanged. A
combined type-plus-text edit therefore remains on the existing fail-closed
cached-text path.

The edit-plan proof accepts the graphic-frame table owner, re-reads the same
physical cell and field ordinal from the source, checks the expected token,
and replaces only the direct `type` attribute on that field's start tag. It
does not rebuild the table or text body.

## Risks / Trade-offs

- [Field ordinal drift] → Count automatic and static fields in source order,
  and bind both cell and field ordinals into the native leaf ID and edit plan.
- [Accidental table flattening] → Require the existing rectangular table and
  fixed text-body profile; reject missing or ambiguous field markup.
- [Semantic overclaim] → Keep automatic field evaluation and host refresh
  source-owned and test only package structure, changed-part scope, and
  re-projection.
