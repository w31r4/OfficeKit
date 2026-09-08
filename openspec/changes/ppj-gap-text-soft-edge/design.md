# Design: authored text soft edge

## Scope

- `textStyle.softEdge` applies wherever the PPJ text-style object is used:
  paragraph `defaultText`, per-run `style`, named text styles, and authored
  master/layout text defaults.
- PPJ owns one `radius` in the existing `0..1000` point profile.
- Native lowering stores the radius as DrawingML EMUs.
- A text run or default run may combine soft edge with bounded glow, inner
  shadow, outer shadow, and reflection. The deterministic order is `a:glow`,
  `a:innerShdw`, `a:outerShdw`, `a:reflection`, then `a:softEdge`.

## Boundaries

This is an authored source-free field and uses the embedded PPJ snapshot for
exact declarative recovery. Imported/source-bound text effects and effect
lists with unsupported children remain source-owned; the codec does not
flatten them into a text soft edge or issue a source-bound text-effect
capability.
