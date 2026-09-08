## Context

See `proposal.md` for the motivation. The custom-geometry importer already
models `a:ahPolar` radial and angle adjustments as literal values or
references, and retains handle identity by ordered index and controlled guide
names. Generic PPJ native leaves are source-bound scalar operations proved
against the original XML before a token splice.

## Goals / Non-Goals

**Goals:**

- Give a direct literal `a:ahPolar/@maxR` bound a stable native index and
  explicit shape-local EMU value.
- Keep radial/angle guide identity, `minR`, angle bounds, position, handle
  order/kind, and custom-geometry topology source-bound.
- Prove a SlidePart-only radial-bound edit and second projection with a valid
  existing range.

**Non-Goals:**

- Editing `minR`, angle bounds, polar position, guide identity, adjustment
  values, formulas, XY handles, paths, or list topology.
- Evaluating or rewriting guide formulas or inferring a new radial range.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use the ordered handle index and polar kind as identity.** This matches
   the existing native handle contract and includes any intervening XY handles.

2. **Require a canonical non-negative integer.** Polar radius is a bounded
   DrawingML integer in shape-local coordinates; exact round-tripping rejects
   references and non-canonical spellings while retaining source-owned state.

3. **Keep the paired range graph intact.** The writer proves all direct
   `ahPolar` attributes and changes only `maxR`; it does not add/remove the
   range or edit `minR`, angle state, current adjustment, or position. The
   focused fixture chooses a replacement that remains inside the existing
   radial range.

4. **Use the generic native-leaf contract.** The new numeric kind relies on
   already modeled Open XML attributes and needs no protobuf descriptor or
   Office wire change.

## Risks / Trade-offs

- An invalid radial min/max/current-adjustment relationship could affect host
  rendering. Restrict the regression to a valid replacement and preserve all
  other range tokens; full range validation remains outside this leaf.
- Changing a radial bound may alter how a host interprets a polar drag handle.
  The operation is scalar source-bound and does not promise host UI behavior or
  automatic path recalculation.
- Legal but unmodeled extension attributes may be present. Reject unknown
  attributes/children and keep the source opaque.

