## Why

F-07 chart text still rejects native kern attributes, although ordinary PPJ text already exposes kerning. This prevents otherwise supported chart styles from retaining that direct setting during semantic edits.

## What Changes

- Add chartTextStyle.kerning using ordinary text's finite 0 through 768 point range.
- Preserve native hundredth-point precision and explicit zero independently of omission.
- Propagate through chart/rich-label styles, field precedence, native edits and vector text without changing literal characters.
- Verify creation/change/zero/deletion/recreation using existing line/combo and vector fixtures.

## Capabilities

### New Capabilities
- `ppj-chart-text-kerning`: Direct chart kerning threshold and its editing lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema, additive wire field, shared chart style codec, authored/source-bound/projector/vector mapping, minimal regressions, chart guidance and generated metadata. No new dependencies or host evaluation workflow.
