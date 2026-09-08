# Design: authored soft-edge effect

## Scope

- `design.styles.shape[].style.softEdge` applies to authored shapes and lines
  through the existing shape-style owner.
- `design.styles.image[].style.softEdge` and direct image `softEdge` apply to
  authored pictures.
- PPJ radius is in points and is bounded to `0..1000`.
- Native wire radius is EMUs and lowers to one direct `a:softEdge` child.
- When authored glow and/or outer shadow are present, the canonical order is
  `a:glow`, `a:outerShdw`, `a:softEdge`.

## Boundaries

The compiler owns only the bounded source-free profile. Imported reflection,
inner-shadow, soft-edge, 3-D, extension-bearing, and arbitrary multi-effect
graphs remain source-preserved or fail closed. This slice does not add a
source-bound capability or infer a soft edge from an imported effect list.

The wire field is optional for compatibility, but authored PPJ requires a
radius. Existing unsupported effect children reject authored composition
instead of being reordered or discarded.
