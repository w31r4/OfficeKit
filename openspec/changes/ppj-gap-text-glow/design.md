# Design: authored text glow

## Scope

- `textStyle.glow` applies wherever the PPJ text-style object is used:
  paragraph `defaultText`, per-run `style`, named text styles, and authored
  master/layout text defaults.
- PPJ owns the existing glow fields: direct RGB or theme color, `radius` in
  points, and optional numeric or opacity-token `opacity`.
- Native lowering stores radius as DrawingML EMUs and opacity as
  thousandth-percent alpha.
- A text run or default run may combine glow with the existing bounded outer
  shadow. The deterministic order is `a:glow`, then `a:outerShdw`.

## Boundaries

This is an authored source-free field and uses the embedded PPJ snapshot for
exact declarative recovery. Imported/source-bound text effects and effect
lists with unsupported children remain source-owned; the codec does not
flatten them into a text glow or issue a source-bound text-effect capability.
