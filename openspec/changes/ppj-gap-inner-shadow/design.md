# Design: authored inner-shadow effect

## Scope

- `design.styles.shape[].style.innerShadow` applies through the existing shape
  style owner.
- `design.styles.image[].style.innerShadow` and direct image `innerShadow`
  apply to authored pictures.
- PPJ owns `color`, `blur`, `distance`, `angle`, and optional `opacity` using
  the same bounded units and opacity grammar token as the existing shadow.
- Native wire values use EMUs, 1/60000 degrees, and thousandth-percent alpha.
- Canonical effect order is `a:glow`, `a:innerShdw`, `a:outerShdw`,
  `a:softEdge` when those bounded authored effects are present.

## Boundaries

Only source-free authored lowering is included. Imported inner shadows and
arbitrary effect lists remain source-preserved or fail closed. The codec does
not infer an inner shadow from a native effect graph or issue a source-bound
editing capability.

Existing unsupported effect children reject authored composition rather than
being silently removed or reordered.
