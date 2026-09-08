## Context

See `proposal.md` for the motivation. The existing custom-geometry reader
already recognizes ordered `a:cxnLst/a:cxn/a:pos` entries and retains literal
or reference coordinates. The x-coordinate leaf establishes the source-bound
scalar contract for this same position element.

## Goals / Non-Goals

**Goals:**

- Give a direct literal y coordinate a stable native index and explicit EMU
  value.
- Reuse the existing native-leaf validation and SlidePart-only scalar splice
  contract used by the angle and x leaves.
- Prove that a changed y value survives a second projection while adjacent
  connection-site and custom-geometry data remains intact.

**Non-Goals:**

- Editing x coordinates, angles, formulas, handles, paths, or list topology.
- Evaluating a guide/reference or recalculating descendants after the edit.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use one leaf per ordered site index.** The index is the existing native
   identity for connection-site targeting. A pair-level position object was
   rejected because this change should authorize only one XML attribute.

2. **Require a canonical non-negative integer token.** Resolved model values
   cannot prove that a source token was literal. Exact integer round-tripping,
   non-negative bounds, and the shape-height bound keep source preconditions
   meaningful; references and formula evaluation remain outside the leaf.

3. **Locate the direct position and splice `@y`.** The writer proves one direct
   custom geometry, one direct connection-site list, one direct position child
   per site, and only the allowed `ang`, `x`, and `y` attributes. It changes the
   selected position's y attribute while leaving the remaining XML sequence
   and package relationships alone.

4. **Use the generic native-leaf contract.** The new kind is a numeric PPJ
   native leaf backed by the already modeled connection-site coordinate, so no
   protobuf descriptor or Office wire version change is needed.

## Risks / Trade-offs

- [Risk] The model may hide whether a coordinate came from a literal or a
  reference. → Require the raw source token to be canonical and re-prove it
  during edit-plan validation and readback.
- [Risk] A legal source geometry may carry unmodeled extension content. →
  Reject non-direct or extra structure for this leaf and preserve the source
  as opaque.
- [Risk] A y change can influence host connector routing. → Keep the edit
  scalar and source-bound; do not promise automatic rerouting or full host
  rendering equivalence.
