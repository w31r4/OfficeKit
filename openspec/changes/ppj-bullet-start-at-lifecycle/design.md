## Context

See proposal.md. The schema and authored compiler already lower `startAt` into optional `PresentationAutoNumberBullet.StartAt`; projection already reads it. The ordinary source compiler masks independent paragraph fields but has no start-value mask/capability. The shared bullet codec rebuilds a modeled marker when applying paragraph properties and reads `StartAt.Value` directly, risking lexical loss and exceptions on malformed tokens.

## Goals / Non-Goals

**Goals:** Retain one numbered marker's existing identity and update only its optional start attribute, with the established text/shape source lifecycle.

**Non-Goals:** Whole-marker edits, automatic numbering evaluation, relationship changes and broader list inheritance. Existing authored/table/master paths remain compatible.

## Decisions

- Add an independent nested `bullet.startAt` mask and exact capability. Lower only paragraphs whose start field changes, leaving marker kind/scheme and other wire fields untouched. Masking the whole bullet would grant unrelated edits and is unsuitable. The added field makes the emitted paragraph capability contain 33 entries, so raise the generic fields-array bound from 32 to 64 while retaining unique entries and the closed field/operation vocabulary.
- When the existing and requested marker are automatic with the same scheme, retain the element and update only the presence/value of `startAt`. Avoid rewriting equal values so noncanonical but valid source spelling survives unrelated edits.
- Parse the raw optional start attribute with invariant integer parsing and explicit 1–32767 bounds. Invalid values leave the marker unmodeled rather than throwing from the SDK typed accessor.
- Reuse the existing source lifecycle helpers, asserting the exact non-target XML/ZIP footprint. Test text and shape owners, mixed styles and a source-owned invalid marker. Add a small automatic-number diagnostic assertion; no new preview implementation.

## Risks / Trade-offs

- Shared bullet application also serves table/list/master owners → retain replacement behavior for different marker schemes/types and run the narrow existing bullet/list regressions.
- Zero or absence can be confused by default-value handling → test explicit 1, absent start and restoration through fresh native projection.
- Unknown XML can disappear during replacement → preserve the original same-scheme element, inject unknown attributes/children and compare the source after removing only the targeted start attribute.
- Broader F-03 remains incomplete → document this field increment without promoting the whole gap or claiming host verification.
