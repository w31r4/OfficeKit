# Design

## Contract

- `glow` is available on `shapeStyle`, `imageStyle`, and the direct image
  element effect surface.
- The field requires `color` and a `radius` in points from `0..1000`.
- `opacity` is optional and accepts a numeric value or an opacity grammar
  token, matching the existing shadow contract.
- The native wire stores the radius in EMUs and opacity in DrawingML
  thousandths of a percent. RGB and theme colors remain mutually exclusive.
- Authored output writes one direct `a:glow`; when an outer shadow is also
  present, the writer keeps the schema order `a:glow` before `a:outerShdw`.

## Boundaries

This is an authored/source-free lowering slice. Imported effect lists with
glow or any other effect are not opened by this change and remain
source-preserved or fail closed. The field does not model reflection, inner
shadow, soft edge, 3-D, SVG filters, effect schemes, or host-specific glow
rendering.
