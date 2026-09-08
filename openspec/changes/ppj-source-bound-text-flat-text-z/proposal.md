# Proposal: expose the bounded flat-text Z coordinate

The PPJ text-body profile now preserves text-warp presets and literal
adjustments, but a direct `a:flatTx/@z` coordinate still disappears from the
semantic text style. Add one bounded integer field and one native leaf while
leaving the surrounding 3D scene source-owned.

## What Changes

- Add `textBoxStyle.flatTextZ` for a signed bounded `a:flatTx/@z` coordinate.
- Author and project one direct `a:flatTx` leaf with a canonical signed integer
  `z` attribute.
- Allow an indexed-free `textBodyFlatTextZ` native leaf to replace only the
  existing `z` token in the owning SlidePart.
- Keep missing/duplicate/extra-attribute/noncanonical/out-of-range `flatTx`
  markup source-owned or fail closed; this does not model `scene3d`, `sp3d`,
  camera, lighting, extrusion, bevel, or host rendering.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-flat-text-z`: bounded authored and source-bound
  preservation of the direct flat-text Z coordinate.

## Impact

- Presentation protobuf, generated bindings, PPJ schema, registry and native
  leaf reference.
- Native body-property projection, authored compilation and token-splice edit
  plan.
- One focused authored/imported/source-bound/reprojection regression.
