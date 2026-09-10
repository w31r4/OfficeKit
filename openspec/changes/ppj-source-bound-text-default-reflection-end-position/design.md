## Context

The native reflection message already preserves optional `stPos` and `endPos`
values. Paragraph default-text projection already reads direct reflection
attributes, while the native leaf projector and edit-plan path currently issue
only the other default reflection scalars and the newly bounded start-position
leaf.

## Goals / Non-Goals

**Goals:**

- Issue one paragraph-index-bound numeric leaf for an existing canonical
  `endPos` while retaining the direct owner and source hash proof.
- Permit that leaf only when `startPos` is absent or `0`, and splice the
  selected XML attribute without rebuilding the effect list.
- Preserve strict handling of unknown children, duplicate owners, and stale
  proofs.

**Non-Goals:**

- Do not add or change protobuf fields or the protocol version.
- Do not expose a reflection with two variable endpoints, WordArt, or host
  PowerPoint rendering behavior in this change.

## Decisions

Use `PresentationReflection.EndPositionThousandthPercent` and represent its
PPJ semantic value as a 0..1 number. The native binding keeps the canonical
0..100000 token as the proof value. Add a dedicated leaf kind and position
validator rather than treating the token as opacity.

Keep the existing safe reflection reader for all other profiles. Extend its
paragraph-default helper with an opt-in end-position allowance that still
requires `startPos` to be absent or `0`; this avoids widening unrelated
reflection edits and rejects a two-variable ramp.

The focused regression authors a reflection with `stPos=0` and `endPos=80000`,
removes the embedded PPJ, projects it, changes only the end-position leaf, and
verifies SlidePart-only output, valid Open XML, preservation of `stPos` and
other attributes, and a second projection. A negative fixture with both
endpoints variable remains opaque.

## Risks / Trade-offs

- [Risk] A relaxed reader could admit a two-variable ramp → [Mitigation] the
  opt-in proof path explicitly checks absent or raw `0` start position.
- [Risk] A numerically valid token may be noncanonical → [Mitigation] compare
  the exact decimal token and validate the 0..100000 integer range before
  patching.
