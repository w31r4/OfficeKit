## Context

See `proposal.md` for the motivation. The existing custom-geometry profile
recognizes ordered `a:ahXY` and `a:ahPolar` handles and validates their
positions against the shape frame. The x-coordinate leaf establishes the
source-bound scalar contract for this same `a:ahXY/a:pos` element.

## Goals / Non-Goals

**Goals:**

- Give a direct literal y position on an XY handle a stable native index and
  explicit EMU value.
- Keep `gdRefX/gdRefY`, bounds, x, handle kind/order, and all custom geometry
  topology source-bound.
- Prove a SlidePart-only y edit and second projection with a focused fixture.

**Non-Goals:**

- Editing handle guide identity, min/max bounds, x position, polar handles,
  adjustment values, formulas, paths, or list topology.
- Evaluating formulas or rerouting connectors after the handle moves.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use the ordered handle index as identity.** The direct `ahLst` order is
   already retained by the native model; matching by coordinates or guide
   names could be ambiguous and could alter identity.

2. **Limit the leaf to direct XY literal y positions.** XY and polar handles
   have different semantics, so this leaf authorizes only `ahXY/pos/@y`.
   Canonical literal parsing and the shape-height bound prevent a resolved
   reference from being mistaken for an editable source token.

3. **Splice the nested position attribute.** The writer proves one direct
   custom geometry, one direct `ahLst`, structurally valid handle entries, and
   one direct `a:pos` under the selected `ahXY`; it then replaces only `@y`.
   Guide references, range attributes, x, and package relationships remain
   untouched.

4. **Use the generic native-leaf contract.** The new numeric kind relies on
   the already modeled handle position, so no protobuf descriptor or Office
   wire version change is needed.

## Risks / Trade-offs

- [Risk] Moving a handle can alter the rendered path after host evaluation. →
  Keep this a scalar source-bound edit and do not promise connector rerouting
  or full host rendering equivalence.
- [Risk] A handle identity change could silently retarget an adjustment. →
  Do not expose `gdRefX/gdRefY` as editable values; prove the selected XY kind
  and retain all guide/range attributes.
- [Risk] Legal but unmodeled handle extensions may be present. → Reject
  unknown attributes/children and preserve the source as opaque.
