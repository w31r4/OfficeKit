# Context

The existing 3-D reader proves direct root attributes and a single child-free
bevel/color owner before exposing scalar leaves. Theme colors use the same
strict owner boundary as direct RGB, but the color child is a bare recognized
`a:schemeClr` rather than `a:srgbClr`.

# Goals / Non-Goals

**Goals:**

- Expose one canonical supported DrawingML theme token for 3-D contour color.
- Replace only the direct `schemeClr/@val` token in the owning SlidePart.
- Support both ordinary shape and picture owners through the same native leaf.
- Preserve all other 3-D attributes/children, image relationships, effects, and
  unrelated package parts.

**Non-Goals:**

- Alpha or color transforms on the 3-D color child.
- Source-free authoring or reconstruction of the surrounding 3-D graph.
- Extrusion color, complex effect/extension topology, or shared relationships.

# Decisions

- Reuse the strict root/owner topology of the existing 3-D RGB reader and add a
  scheme-specific color reader that accepts only the canonical theme vocabulary.
- Store the projected token in additive shape/image fields and reject those
  fields during new-object authoring because they are source-bound state.
- Use the existing generic shape/picture native-leaf proof and XML splice path,
  changing only the color element name and its `val` attribute.

# Risks / Trade-offs

- A theme color with transforms could be mistaken for a plain token. → Require
  a child-free `a:schemeClr` and reject all transforms/extensions.
- Shape and picture paths could diverge. → Exercise both owners in focused
  tests and keep the native leaf shared.
