## Context

Custom geometry text rectangles are represented by `a:rect` edge tokens. The
OfficeKit authoring path uses four private `officeKitText*` scaled guides for
literal edges, while imported files may use direct numeric tokens or named
guide references. Generic PPJ native leaves are source-bound scalar operations
proved against the original XML before a token splice.

## Goals / Non-Goals

**Goals:**

- Give a fully literal, in-frame text-rectangle top edge a stable native leaf
  and shape-local EMU value.
- Preserve the other three edges, the private guide/rect representation,
  guide identity, paths, and all other custom-geometry topology.
- Prove the authored private-guide profile and a direct numeric source with a
  SlidePart-only edit and a second projection.

**Non-Goals:**

- Editing left/right/bottom edges, guide references, guide formulas, handles,
  connection sites, paths, or list topology.
- Converting arbitrary guide references to literals or changing the private
  guide profile shape.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use one native index for the top edge.** The leaf represents the one
   recognized `a:rect` owner on the shape; its value is not a second rectangle
   vector.

2. **Accept only canonical in-frame literal profiles.** Direct numeric edge
   tokens and the four exact OfficeKit scaled guides are allowed. Every edge
   must be non-negative and within its shape-local extent, and left < right and
   top < bottom must hold before and after the edit.

3. **Preserve representation.** A direct source keeps `a:rect/@t` numeric. An
   authored/private source keeps `t="officeKitTextTop"` and changes only the
   matching `fmla` numeric operand, preserving the scale axis and denominator.

4. **Use the generic native-leaf contract.** The new numeric kind needs no
   protobuf descriptor or Office wire change.

## Risks / Trade-offs

- A top-edge change can alter text layout inside a custom shape. Restrict the
  regression to an ordered in-frame change; reflow behavior remains host-owned.
- A source may contain a legal text rectangle with negative or formula-backed
  edges. Those remain source-owned rather than being forced into this bounded
  profile.
- Legal but unmodeled extension attributes may be present. Reject unknown
  rectangle/guide structure and keep the source opaque.
