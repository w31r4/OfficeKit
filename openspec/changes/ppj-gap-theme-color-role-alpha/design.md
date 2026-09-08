# Design

## Contract

- `design.theme.accentColors.accent1..accent6` and
  `design.theme.colorRoles.dark1/light1/dark2/light2/hyperlink/followedHyperlink`
  accept the existing `#RRGGBB` form and the new `#RRGGBBAA` form.
- The last two hex digits use ordinary RGBA meaning: `00` is transparent and
  `FF` is opaque.
- The compiler writes the six-digit RGB value to `a:srgbClr/@val` and, when
  the alpha suffix is present, writes one `a:alpha/@val` in the 0..100000
  DrawingML scale.

## Ownership and boundary

The source-free compiler owns the complete bounded theme it creates. An
imported theme remains source-owned; adding or changing these authored theme
fields on a source-bound PPJ remains rejected. No theme projection is changed
to infer transforms, effect styles, or host font behavior.

## Evidence

Use a minimal source-free PPJ fixture with one translucent accent and one
translucent non-accent role. Assert schema validation, `theme1.xml` values,
Open XML validation, and preservation of the exact embedded PPJ bytes.
