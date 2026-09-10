## Context

See proposal.md. The wire already has point/multiplier/no-space-before alternatives. The native spacing codec supports values and deletion, but PPJ lacks field authority and deletion lowering. Its schema bounds exceed the native range; the authored builder selects the two units separately, allowing lower-priority multipliers to override direct points.

## Goals / Non-Goals

Complete the direct space-before owner in both units. Existing space-after, line-spacing, table and inherited-placeholder editing boundaries remain separate work.

## Decisions

Retain both existing PPJ numeric fields and reject simultaneous declarations. Select the highest-priority style containing either field, and normalize its value before semantic comparison. Reuse the wire's no-space-before intent for deletion.

Add one combined change mask, then require authority for every changed unit field. Lower only changed paragraphs. Read only a unique bounded native slot, using raw numeric attributes and rejecting hidden descendants. Unmodeled space before can remain in an otherwise editable paragraph because its replacement is checked locally; unrelated changes leave it untouched.

Skip native replacement when unit/value are unchanged to preserve source spelling in neighboring and other spacing slots. Replace an existing target in place, and insert a new space-before node after line spacing in native order. This avoids rewriting a paragraph or its spacing siblings.

Use the existing original-source XML/ZIP harness with small lifecycle, style-precedence and rejected-source fixtures. The generic unsupported-edit fixture moves to spaceAfter now that spaceBefore becomes editable.

## Risks / Trade-offs

Zero is confused with absence or unit → compare native child type and fresh PPJ for both units.

Authored style precedence or rounding differs from source editing → test direct overrides and native precision.

Permitting unrelated edits with unmodeled space before enables replacement accidentally → test local rejection plus unrelated scalar preservation.

Partial preview is mistaken for host layout evidence → keep its diagnostic and record unperformed NativeAOT/host checks.
