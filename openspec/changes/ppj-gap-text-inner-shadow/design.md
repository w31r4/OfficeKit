# Design: authored text inner shadow

## Scope

- `textStyle.innerShadow` applies wherever the PPJ text-style object is used:
  paragraph `defaultText`, per-run `style`, named text styles, and authored
  master/layout text defaults.
- PPJ owns the existing inner-shadow fields: direct RGB or theme color,
  `blur`, `distance`, and `angle` in points/degrees, plus optional numeric or
  opacity-token `opacity`.
- Native lowering stores point distances as DrawingML EMUs, direction as the
  normalized 60000ths-of-a-degree value, and opacity as thousandth-percent
  alpha.
- A text run or default run may combine inner shadow with the bounded glow and
  outer shadow. The deterministic order is `a:glow`, `a:innerShdw`, then
  `a:outerShdw`.

## Boundaries

This is an authored source-free field and uses the embedded PPJ snapshot for
exact declarative recovery. Imported/source-bound text effects and effect
lists with unsupported children remain source-owned; the codec does not
flatten them into a text inner shadow or issue a source-bound text-effect
capability.
