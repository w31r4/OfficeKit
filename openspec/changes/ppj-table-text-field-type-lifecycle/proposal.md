## Why

Fixed-topology table cells already expose typed fields and allow cached display
text edits, but their DrawingML field type token remains source-owned. This
leaves the table-field portion of the F-06 P1 gap unable to express a safe
static field-type change.

## What Changes

- expose `table.rows[].cells[].text.paragraphs[].runs[].field.type` as a
  source-bound native field for eligible rectangular table cells;
- add a `setTextField` capability entry for the table-cell path;
- accept only changed valid non-automatic type tokens and patch the direct
  `a:fld/@type` attribute in the owning slide;
- preserve field ID, cached text, automatic state, paragraph/run topology,
  table grid, merges, and all non-target XML; reject identity, automatic,
  topology, stale-authority, and unsupported-field changes;
- add a focused authored/imported round-trip regression and update PPJ
  schema, registry, Skill reference, coverage, and the F-06 backlog.

## Capabilities

### New Capabilities

- `ppj-table-text-field-type-lifecycle`: source-bound static field-type
  projection and token-only editing for fixed-topology table cells.

### Modified Capabilities

<!-- No existing OpenSpec capability requirements are present in this repo. -->

## Impact

The native PPJ table projection, source-bound compiler, edit-plan proof and
XML patcher gain one table-field leaf profile. The public PPJ schema and
capability registry gain the table-cell field path; no protobuf wire version
change is required. Unsupported automatic fields, non-rectangular tables,
and complex field graphs remain source-preserved or fail closed.
