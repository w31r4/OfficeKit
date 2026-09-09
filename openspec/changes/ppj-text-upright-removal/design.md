## Context

See proposal.md. MergeSourceBoundTextBodyStyle only applies present fields. Tables rebuild structured text separately. Native BodyProperties already has an explicit no_upright deletion operation and the writer removes upright when selected.

## Goals / Non-Goals

Complete upright true/false/absence semantics across existing editable text owners. Keep other body-style omission semantics and unsupported/inherited owner boundaries unchanged; no visual layout or host acceptance is added.

## Decisions

Pass previous source style into the merge so a previously present upright becoming absent selects native NoUpright. Explicit false remains a value. Permit removing the style object only when its sole prior property was upright; keep ordinary whole-style removal guards. Structured table text uses its old/new style when building the requested body and applies the same deletion marker after validating the represented body. Allow that valid deletion marker through the bounded native table body check. Reuse existing source capabilities and writer, avoiding new schema null syntax or wire state.

Removing the last table-body property can make the native reader select its existing compact text/cell-style representation. Share that uniform-body predicate with semantic hash normalization after clearing edit intent, so post-write comparison uses the reader's representation without discarding nonuniform runs or remaining body/paragraph state. Verify the resulting plain cell can regain upright through structured text with the same native topology.

## Risks / Trade-offs

Silent deletion of unrelated style -> compare only upright and retain whole-style guard. Table path discrepancy -> real table source regression. No-op rewrites -> byte identity assertion. Native marker leaking into projection -> fresh projection asserts property omission. Unsupported owners -> existing capability checks remain in force.

## Migration Plan

Behavior fix for source PPJ snapshots: delete upright to remove the native attribute. Authored schema remains boolean with omission. Update descriptions and focused evidence alongside implementation.
