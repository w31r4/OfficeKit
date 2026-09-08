# Design: authored text reflection

## Scope

- `textStyle.reflection` applies wherever the PPJ text-style object is used:
  paragraph `defaultText`, per-run `style`, named text styles, and authored
  master/layout text defaults.
- PPJ owns the existing reflection fields: `blur`, `startOpacity`,
  `endOpacity`, `distance`, and `angle` in the bounded points/degrees profile.
- Native lowering stores point distances as DrawingML EMUs, direction as the
  normalized 60000ths-of-a-degree value, opacity as thousandth-percent alpha,
  and writes the deterministic full-span `stPos="0"`/`endPos="100000"`.
- A text run or default run may combine reflection with the bounded glow,
  inner shadow, and outer shadow owners already available to text. The
  deterministic text order is `a:glow`, `a:innerShdw`, `a:outerShdw`, then
  `a:reflection`.

## Boundaries

This is an authored source-free field and uses the embedded PPJ snapshot for
exact declarative recovery. Imported/source-bound text effects and effect
lists with unsupported children remain source-owned; the codec does not
flatten them into a text reflection or issue a source-bound text-effect
capability.
