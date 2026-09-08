# Proposal: expose a bounded text-warp preset

The PPJ text-body profile currently preserves the WordArt marker but leaves
the direct DrawingML `a:prstTxWarp` preset opaque. Exposing the preset name
closes one independently testable part of the WordArt gap while keeping its
adjustment formula list and full rendering behavior source-owned.

## What Changes

- Add `textBoxStyle.textWarpPreset` to authored and projected PPJ.
- Preserve the finite DrawingML text-shape preset vocabulary in the additive
  presentation wire model and issue `textBodyWarpPreset` for a safe source
  leaf.
- Support authored output and source-bound replacement of one existing
  `a:prstTxWarp/@prst` value when the element is bare or has a recognized
  literal adjustment list.
- Keep non-literal adjustment values, extensions, and the rest of the
  WordArt transform graph opaque or fail closed.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-warp-preset`: bounded authored and source-bound
  preservation of the direct text-warp preset.

### Modified Capabilities

## Impact

- Presentation protobuf, generated JavaScript bindings, PPJ schema and
  capability registry.
- Native body-property projection, authored compilation, edit-plan proof and
  XML token splicing.
- One focused NativeAOT codec regression and synchronized presentation Skill
  reference.
