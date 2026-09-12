# Design

## Context

See `proposal.md` and the direct static field-type lifecycle for the existing
PPJ field path, field ordinal, source hash, and source-bound text proof. The
codec already preserves imported package bytes and can splice one attribute
token in an owning slide.

## Goals / Non-Goals

**Goals:**

- Issue a strict `textFieldId` leaf only for direct, non-automatic fields with
  a valid existing ID.
- Re-prove the same field and replace only the `a:fld/@id` value.
- Make the result observable through a second projection and a focused
  authored/imported regression.

**Non-Goals:**

- Changing field type, cached display text, automatic evaluation, refresh
  behavior, table-field identities, or relationship graphs.
- Rebuilding a text body or normalizing unrelated XML.

## Decisions

The public leaf uses the existing `native-leaf` surface and the existing field
ordinal. This keeps the PPJ shape stable and lets the compiler reuse the
source-bound identity proof used for `textFieldType`; a separate object or
relationship model would widen the contract without evidence that PowerPoint
consumes it.

IDs are validated with the codec's existing brace-wrapped UUID grammar. The
compiler compares the exact source token before replacing only the attribute
value, so whitespace, prefixes, cached text, and neighboring XML remain
untouched. Automatic fields do not receive the leaf even when their XML has an
ID, because their identity is tied to host evaluation. Table field IDs stay
opaque until a table-specific profile exists.

## Risks / Trade-offs

- [External references may assign meaning to a field ID] → limit the edit to
  direct static fields and document that relationship semantics remain
  source-owned.
- [A malformed imported ID cannot be safely reprojected] → fail closed and do
  not issue an editable leaf.
- [The focused test does not prove host PowerPoint behavior] → report the
  evidence boundary explicitly and avoid a host-acceptance claim.

## Migration Plan

No data migration or wire-version change is needed. Existing PPJ documents
without `textFieldId` remain valid; callers can use the leaf only when a fresh
projection advertises it. Reverting the change removes the edit capability but
does not alter source packages.
