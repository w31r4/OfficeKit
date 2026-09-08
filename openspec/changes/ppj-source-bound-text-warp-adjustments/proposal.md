# Proposal: expose bounded text-warp adjustments

The PPJ text-body profile now preserves a finite `a:prstTxWarp/@prst` preset,
but a preset with literal DrawingML adjustment values still loses the editable
portion of its WordArt shape. This change adds one ordered adjustment field and
one indexed native leaf while keeping formula-driven and extended WordArt
topologies source-owned.

## What Changes

- Add `textBoxStyle.textWarpAdjustments` as an ordered list of `{name, value}`
  entries paired with the existing `textWarpPreset` field.
- Accept only one direct `a:prstTxWarp` whose optional `a:avLst` contains
  unique, attribute-only `a:gd name="..." fmla="val N"` entries.
- Preserve authored values, project them after source-bound import, and allow
  one indexed `textBodyWarpAdjustment` leaf to replace only a selected `val N`
  token.
- Keep non-literal formulas, extension children, duplicate names, and other
  WordArt transform/effect state source-owned or fail closed.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-warp-adjustments`: bounded authored and source-bound
  preservation of literal `a:prstTxWarp` adjustment values.

### Modified Capabilities

- `ppj-source-bound-text-warp-preset`: the existing preset field may now be
  projected alongside a recognized literal adjustment list; malformed or
  unsupported lists still suppress the bounded preset leaf.

## Impact

- Presentation protobuf, generated JavaScript bindings, PPJ schema and
  capability registry.
- Native body-property projection, authored compilation, edit-plan proof and
  XML token splicing.
- One focused regression for authored/imported/source-bound adjustment values
  and fail-closed formula/extension cases, plus synchronized Skill coverage.
