## Context

See proposal.md. Existing authored/default-run readers use FontKerningPoints; schema and validation bound it to 0..768. The native attribute uses hundredths of a point. Source paragraph masks currently omit kerning authority.

## Goals / Non-Goals

Complete the direct default threshold lifecycle for ordinary text/shape paragraphs. Run kerning, placeholder inheritance, table defaults and host font shaping retain their separate profiles.

## Decisions

Add the exact field to projection/schema/semantic authority and paragraph diff masks. Clone the existing default style, then clear/copy only kerning presence. Canonicalize finite in-range requests to native hundredths using ties-to-even rounding before postwrite comparison, following defaultText.size.

Extend the existing targeted default scalar writer instead of rebuilding the whole default style, preserving unrelated effects and unknown XML. Patch only changed kerning; use direct null assignment for deletion. Reject replacing native kern outside the modeled range. Clear only modeled kerning when removing a default style, so unrelated field deletion preserves unmodeled source values.

## Risks / Trade-offs

Explicit zero and omission can collapse → compare native attribute and fresh PPJ presence in the shared fixture. Rounding can cause postwrite mismatch → canonicalize requested semantics to the actual hundredth. Unknown native kern can be discarded by cleanup → one fixture exercises no-op, replacement rejection, unrelated assignment and removal. Host kerning appearance remains unverified.
