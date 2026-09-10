## Context

The native reflection message already preserves optional `stPos` and `endPos`
values, and the paragraph default-text reader accepts bounded direct reflection
attributes. The native leaf projector and edit-plan path currently issue only
the other default reflection scalars and require full-span positions for their
proof.

## Goals / Non-Goals

**Goals:**

- Issue one paragraph-index-bound numeric leaf for an existing canonical
  `stPos` while retaining the existing direct owner and source hash proof.
- Permit that leaf only when `endPos` is absent or `100000`, and splice the
  selected XML attribute without rebuilding the effect list.
- Preserve the existing strict handling of unknown children, duplicate owners,
  and stale proofs.

**Non-Goals:**

- Do not add or change protobuf fields or the protocol version.
- Do not expose `endPos`, variable reflection transforms, WordArt, or host
  PowerPoint rendering behavior in this change.

## Decisions

Use the existing `PresentationReflection.StartPositionThousandthPercent` field
and represent its PPJ semantic value as a 0..1 number while retaining the raw
100000-based token in `nativeRef.leaves[].value`. Add a dedicated leaf kind and
position validator rather than treating the token as opacity, so the field's
meaning and diagnostics stay explicit.

Keep the existing safe reflection reader for all other profiles. Extend its
paragraph-default helper with an opt-in start-position allowance that still
requires `endPos` to be absent or `100000`; this avoids widening unrelated
reflection edits. The XML patcher maps the new kind to `stPos` and reuses the
paragraph index proof.

The focused regression will author a reflection with full-span positions,
remove the embedded PPJ, project it, change only the start-position leaf, and
verify SlidePart-only output, valid Open XML, preservation of `endPos` and
other attributes, and a second projection. Negative fixtures cover non-full
end positions and malformed/unknown source tokens.

## Risks / Trade-offs

- [Risk] A non-full `endPos` could be accidentally admitted by a relaxed
  reader → [Mitigation] the new proof path explicitly checks absent or raw
  `100000` end position before issuing or accepting the leaf.
- [Risk] A raw position token may be numerically valid but noncanonical →
  [Mitigation] compare the exact decimal token and validate the 0..100000
  integer range before patching.
