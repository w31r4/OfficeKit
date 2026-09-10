## Context

See proposal.md. Space before now has complete unit/presence lowering and source-preserving spacing helpers. Space after still lacks PPJ authority, deletion lowering and unit-aware style precedence; its schema bounds exceed the same native ranges.

## Goals / Non-Goals

Complete the direct space-after slot and preserve the space-before contract. Line-spacing lifecycle, table/placeholder inheritance and host layout retain separate boundaries.

## Decisions

Reuse the existing numeric fields and no-space-after wire intent. Add a combined change mask plus exact authority checks for each changed unit field, and lower only the changed paragraph slot. Extend authored paragraph merging to select space-after units at the highest-priority style.

Read only a unique bounded space-after node. Keep unmodeled content editable only through unrelated state; the local replacement guard rejects overwrites. Reuse raw numeric validation and unchanged-value preservation. Insert a new space-after node after the last line-spacing/space-before predecessor, retaining unknown siblings.

Extract shared helpers from the existing space-before experiment and call them for both slots. Retain its original assertions and add the four equivalent space-after cases without a copied fixture suite. Move the generic unsupported-edit fixture to lineSpacing.

## Risks / Trade-offs

Unit choice or zero disappears during lowering → assert native child type and fresh PPJ, including style precedence.

Shared helper changes disturb space before → run both slot experiments and the affected paragraph/table regressions.

Unknown content is normalized or replaced → retain original XML/ZIP checks and rejected-source cases.

Preview is confused with host layout proof → retain explicit partial diagnostics and separate NativeAOT/host evidence.
