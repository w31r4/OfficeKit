## Context

See proposal.md. PPJ indent maps to native `marL`; hanging maps to the negated native `indent` coordinate. Authored indent already uses ties-to-even EMU rounding. Native left-margin deletion exists, but the PPJ authority/mask/lowering path is missing, and the native reader currently trusts SDK numeric coercion.

## Goals / Non-Goals

Complete ordinary text/shape paragraph left indent with exact field ownership. Hanging lifecycle and table/placeholder inheritance retain separate boundaries.

## Decisions

Reuse `indent` and `no_margin_left`, preserving current naming and avoiding a new wire operation. Add one field mask and authority check, and lower only changed paragraphs. Retain scalar style precedence and native precision.

Parse native left-margin attribute text directly and enforce 0..51206400 EMUs. Move its replacement check from the whole-paragraph support gate into its native setter, so unrelated edits preserve invalid source values. Skip unchanged coordinate assignments to preserve raw left-margin and hanging-indent spelling. Keep hanging validation intact.

Reuse the existing lifecycle harness and XML/ZIP comparison helper for a focused text/shape fixture plus invalid-source cases. Move its generic unsupported paragraph edit from indent to hanging.

## Risks / Trade-offs

Indent is confused with hanging → document and assert independent native attributes.

Zero, native precision or source spelling is lost → assert explicit attributes, nearest-even rounding, deletion/restoration and fresh PPJ.

SDK coercion changes malformed source values → read raw integers and verify no-op, unrelated edits and rejected replacements.

Shared layout behavior regresses → run the narrow fixture and existing paragraph/table regression group; preserve known baseline exclusions.
