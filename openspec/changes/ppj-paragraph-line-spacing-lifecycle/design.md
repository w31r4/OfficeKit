## Context

See proposal.md. Before/after spacing now has unit-aware merging, exact PPJ authority and local native replacement guards. Line spacing still uses independent unit lookups and a paragraph-wide support gate. Its existing native validator requires positive rounded values and already supports no-line-spacing deletion.

## Goals / Non-Goals

Complete direct line spacing with the same source-preservation guarantees as before/after spacing. Table/placeholder inheritance and host line layout retain their separate boundaries.

## Decisions

Reuse existing fields and deletion intent, avoiding a new wire operation. Add a combined change mask and exact per-unit authority checks; lower only changed paragraph slots. Extend paragraph-style merging to choose line-spacing units together, instead of mixing units across priority layers. Keep positive values and reject rounded zero rather than clamping or interpreting zero as deletion.

Read only a unique bounded line-spacing slot. Remove the paragraph-wide spacing gate now that all three slots reject replacement locally, allowing unrelated edits to retain unmodeled source content. Insert a new line-spacing node before its known spacing successors; preserve unknown siblings and unchanged numeric spelling.

Extend the shared before/after experiment to line spacing, with positive minimum and zero-rejection cases. Keep the original XML/ZIP assertions for all three slots. Move the generic unsupported-edit fixture to paragraph indent.

## Risks / Trade-offs

Unit choice or very small values change silently → assert native child type and fresh PPJ, minimum native precision and rounded-zero rejection.

Shared spacing changes disturb existing fields → run all three lifecycle experiments and affected paragraph/table regressions.

Unknown source content is normalized → retain no-op bytes, unrelated-edit XML and rejected replacement checks.

Preview is confused with host proof → keep partial diagnostics and separate NativeAOT/host evidence.
