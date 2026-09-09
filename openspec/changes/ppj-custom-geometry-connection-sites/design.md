## Context

See proposal.md. Native connection sites store optional reference strings or literal EMUs/1-60000 degree angles. Validation resolves the existing adjustment/guide graph, bounds positions to the shape and angles to one signed turn. Native source export fixes list length to preserve connector index identity.

## Goals / Non-Goals

Expose the complete bounded site record for literal-path custom shapes. Keep handle/reference-backed-path owners and masks/clips outside this shape graph profile. Connector endpoint authoring syntax and host dragging remain separate increments.

## Decisions

Use at most 1024 ordered `{angle, x, y}` records, each value a native reference string or number (degrees/local points). Reuse conversion conventions from textRectangle; coordinates are independent of path viewBox. Empty authored lists project as omission. Source setGeometry accepts per-index value edits but rejects length changes, including deletion of a nonempty list, in accordance with native identity. Preserve existing connector binding and rerouting behavior.

## Risks / Trade-offs

Index drift -> retain native fixed-length validation. Formula reference loss -> round-trip mixed literals and references and reject dangling references. Unit confusion -> check concrete native EMUs and angles. Silent preview support -> assert field diagnostics.

## Migration Plan

Additive PPJ schema and setGeometry field vocabulary. No wire changes. Fresh projection recovers the new field.
