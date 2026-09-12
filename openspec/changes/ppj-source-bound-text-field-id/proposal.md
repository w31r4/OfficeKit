# Proposal: add source-bound direct text-field identity

## Why

PPJ already projects a static direct field's type, but its `a:fld/@id` remains
source-owned. That leaves a visible part of the field contract unavailable to
an OfficeKit caller and prevents a small, source-preserving identity edit.

## What Changes

- Expose `text.paragraphs[].runs[].field.id` as a `textFieldId` native leaf for
  valid non-automatic direct text fields.
- Accept a changed brace-wrapped UUID only when the fresh source projection,
  field ordinal, and expected ID still match.
- Token-splice only `a:fld/@id` in the owning slide and preserve the field type,
  cached text, automatic state, inline topology, and all other package parts.
- Keep automatic fields, table fields, field refresh/evaluation, and field
  relationship semantics source-owned.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-field-id`: project and source-bound edit the identity
  of a direct static DrawingML text field.

### Modified Capabilities

None.

## Impact

The PPJ schema, capability registry, generated presentation references and
coverage docs gain one native leaf kind. The native projector and edit-plan
codec gain strict ID validation, source-bound proof, and one-attribute XML
patching. No protobuf or wire-version change is required.
