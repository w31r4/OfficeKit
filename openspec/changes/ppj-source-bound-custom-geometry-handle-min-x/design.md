## Context

See `proposal.md` for the motivation. The existing custom-geometry profile
recognizes XY handle range attributes as literal values or references and
retains handle identity by ordered index and controlled adjustment names.
Generic PPJ native leaves are source-bound scalar operations proved against the
original XML before a token splice.

## Goals / Non-Goals

**Goals:**

- Give a direct literal `a:ahXY/@minX` bound a stable native index and explicit
  EMU value.
- Keep guide identity, the paired max bound, y bounds, position, handle order,
  and custom-geometry topology source-bound.
- Prove a SlidePart-only bound edit and second projection with a valid range.

**Non-Goals:**

- Editing max bounds, y bounds, handle position, guide identity, adjustment
  values, formulas, polar handles, paths, or list topology.
- Evaluating or rewriting guide formulas or inferring a new range pair.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use the ordered handle index and XY kind as identity.** This matches the
   existing native handle contract and avoids retargeting a bound by name or
   coordinate.

2. **Require a canonical non-negative integer.** The resolved model can hide a
   source reference, so exact integer round-tripping and the shape-width bound
   are required. Reference-backed bounds remain opaque.

3. **Keep the paired range graph intact.** The writer proves the direct
   `ahXY` attribute structure and changes only `minX`; it does not add/remove
   the range or edit `maxX`, current adjustment, or other bounds. The focused
   fixture chooses a replacement that remains inside the existing range.

4. **Use the generic native-leaf contract.** The new numeric kind relies on
   already modeled Open XML attributes and needs no protobuf descriptor or
   Office wire change.

## Risks / Trade-offs

- [Risk] An invalid min/max/current-adjustment relationship could affect host
  rendering. → Restrict the regression to a valid replacement and preserve all
  other range tokens; future full range validation remains outside this leaf.
- [Risk] Changing a bound may alter how a host interprets a drag handle. →
  Document that the operation is scalar source-bound and does not promise host
  UI behavior or automatic path recalculation.
- [Risk] Legal but unmodeled extension attributes may be present. → Reject
  unknown attributes/children and keep the source opaque.
