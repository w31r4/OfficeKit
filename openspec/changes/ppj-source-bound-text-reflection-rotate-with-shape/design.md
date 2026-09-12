## Context

`PptxReflectionCodec` already reads and writes `rotWithShape` when reflection transforms are explicitly allowed. Direct rich-text projection and source-bound editing currently reject that flag, even though the paragraph default-text profile exposes the same boolean.

## Goals / Non-Goals

**Goals:**

- Carry an explicit direct-run reflection `rotWithShape` token through PPJ projection and source-bound editing.
- Keep the owner proof strict: full-span direct run, one canonical `0|1` token, no fade, scale, skew, alignment, or ambiguous effect topology.
- Keep the experiment to one authored/source-bound/reprojection regression.

**Non-Goals:**

- Combining rotate-with-shape with other reflection transforms.
- Changing the reflection wire contract or claiming PowerPoint host rendering acceptance.

## Decisions

- Add a boolean `textReflectionRotateWithShape` leaf mapped to `run.style.reflection.rotateWithShape`; preserve explicit false as a leaf because omission and `0` are distinct native states.
- Reuse `PptxReflectionCodec` with transform parsing enabled only for this direct-run profile. A run carrying another transform remains source-owned rather than receiving a partial rotation leaf.
- Reuse the existing text-reflection XML patcher and replace only `rotWithShape`; extend proof, read, validation, and dispatch lists together.
- Document the field in the backlog, coverage, PPJ/text references, capability registry, and generated presentation matrix.

## Risks / Trade-offs

- A source writer may include an unmodeled sibling or noncanonical boolean text. The strict reader and canonical `0|1` precondition keep such owners opaque.
- The shared reflection message is also used by shape/image/chart styles. The new leaf is gated only in the direct-run path, leaving those existing boundaries unchanged.
