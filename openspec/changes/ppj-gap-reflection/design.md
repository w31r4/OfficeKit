# Design: authored reflection effect

## Scope

- `design.styles.shape[].style.reflection` applies through the existing shape
  style owner to authored shapes and lines.
- `design.styles.image[].style.reflection` and direct image `reflection` apply
  to authored pictures.
- PPJ owns `blur` and `distance` in points, `startOpacity` and `endOpacity` as
  numbers or opacity grammar tokens, and `angle` in degrees.
- Native wire values use EMUs, 1/60000 degrees, and thousandth-percent alpha.
- Native lowering writes `stPos="0"` and `endPos="100000"` for a deterministic
  full-span reflection; other DrawingML reflection attributes are outside this
  bounded field.
- Canonical effect order is `a:glow`, `a:innerShdw`, `a:outerShdw`,
  `a:reflection`, `a:softEdge` when those bounded authored effects are present.

## Boundaries

Only source-free authored lowering is included. Imported reflections and
arbitrary effect lists remain source-preserved or fail closed. The codec does
not infer reflection semantics from a native effect graph or issue a
source-bound editing capability.

Existing unsupported effect children reject authored composition rather than
being silently removed or reordered.
