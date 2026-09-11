## Context

`PptxReflectionCodec` already reads and writes DrawingML alignment when transforms are explicitly allowed. Direct rich-text projection and source-bound edit planning currently omit this scalar even though paragraph default-text reflection already exposes the same enum.

## Goals / Non-Goals

**Goals:**

- Carry the nine rectangle-alignment tokens through direct run PPJ projection and source-bound edits.
- Keep the owner proof strict: full-span direct run, one canonical `algn`, no other transform, and no ambiguous effect topology.
- Keep the experiment limited to one authored/source-bound/reprojection regression.

**Non-Goals:**

- Adding rotate-with-shape or combined reflection transform leaves.
- Claiming PowerPoint host rendering or changing the wire protocol.

## Decisions

- Add a string `textReflectionAlignment` leaf and map it to `run.style.reflection.alignment`; use the existing `PptxShadowCodec` enum mapping so authored and imported values share one token set.
- Reuse the general reflection schema and compiler parser, then enable alignment only in the direct-run projection/edit profile. A direct run with any other transform remains source-owned rather than issuing a partial alignment leaf.
- Reuse the existing text-reflection XML patcher and replace only `algn`; extend its proof/read/validation lists in lockstep. This preserves all other effects, attributes, relationships, and ZIP parts.
- Document the field in the backlog, coverage, PPJ reference, capability registry, and generated presentation matrix; regenerate derived skill references through the maintainer script.

## Risks / Trade-offs

- [Risk] A source writer may order attributes differently or include an unmodeled sibling. → Require exactly one canonical attribute and the existing strict reflection reader; otherwise keep the owner opaque.
- [Risk] The shared reflection schema also applies to shape/image styles. → The direct-run leaf is gated separately; existing shape/image projection boundaries are unchanged.
