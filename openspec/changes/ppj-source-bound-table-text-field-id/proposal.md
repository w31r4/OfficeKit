## Why

Fixed-topology table cells already expose cached field text and a static field
type, but the field identity remains source-owned. That leaves one concrete
F-03/F-06 text-field leaf unavailable to PPJ edits even when the surrounding
cell and inline topology are safely bound.

## What Changes

- expose `table.rows[].cells[].text.paragraphs[].runs[].field.id` as a
  `tableTextFieldId` source-bound native leaf for eligible rectangular cells;
- reuse `setTextField` for a changed brace-wrapped UUID and patch only the
  owning direct `a:fld/@id` token;
- preserve field type, cached text, automatic state, inline topology, table
  grid, merges, and non-target XML; reject automatic, invalid, stale,
  combined-text, topology, and unsupported-field edits;
- add a focused authored/imported lifecycle regression and update the PPJ
  schema, registry, Skill reference, coverage, and F-03 backlog.

## Capabilities

### New Capabilities

- `ppj-source-bound-table-text-field-id`: source-preserving identity edits for
  static DrawingML fields inside fixed-topology table cells.

### Modified Capabilities

<!-- No existing OpenSpec capability requirements are present in this repo. -->

## Impact

The table native projection, source-bound presentation compiler, edit-plan
proof, and XML token patcher gain one field-identity leaf. The public PPJ
schema and capability registry gain the table-cell field path. No protobuf
wire change is required; automatic evaluation and field relationships remain
source-owned.
