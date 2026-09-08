## Context

See `proposal.md` for the motivation. The custom-geometry importer already
models `a:ahPolar/a:pos` as a pair of literal coordinates or references, and
retains handle identity by ordered index and controlled guide names. Generic
PPJ native leaves are source-bound scalar operations proved against the
original XML before a token splice.

## Goals / Non-Goals

**Goals:**

- Give a direct literal polar-handle position `y` a stable native index and
  explicit shape-local EMU value.
- Keep guide identity, radial/angular bounds, x, position topology, handle
  order/kind, and custom geometry source-bound.
- Prove a SlidePart-only coordinate edit and second projection with a valid
  existing handle graph.

**Non-Goals:**

- Editing polar position `x`, guide identity, adjustment values, formulas,
  range bounds, XY handles, paths, or list topology.
- Evaluating guide formulas or converting a referenced coordinate into a
  literal.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use the ordered handle index and polar kind as identity.** This matches
   the existing native handle contract and includes any intervening XY handles.

2. **Require a canonical shape-local coordinate.** The direct `@y` token must
   be a non-negative integer no greater than the custom geometry height;
   references and non-canonical spellings remain source-owned.

3. **Keep the entire handle graph intact.** The writer proves all direct
   `ahPolar` attributes and the direct position child, then changes only `y`.
   It does not add/remove the position or edit x, range state, current
   adjustment, or guides.

4. **Use the generic native-leaf contract.** The new numeric kind relies on
   already modeled Open XML attributes and needs no protobuf descriptor or
   Office wire change.

## Risks / Trade-offs

- Moving a polar handle anchor may affect host interpretation of an existing
  custom path. Restrict the regression to an in-frame replacement and preserve
  all other geometry tokens; automatic path recalculation remains outside this
  leaf.
- Legal but unmodeled extension attributes may be present. Reject unknown
  attributes/children and keep the source opaque.
