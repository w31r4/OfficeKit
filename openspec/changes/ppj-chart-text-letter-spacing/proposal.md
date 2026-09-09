## Why

F-07 chart typography cannot yet express direct character spacing, although ordinary PPJ text already has letterSpacing. Imported spc attributes therefore keep otherwise supported chart styles source-owned.

## What Changes

- Add chartTextStyle.letterSpacing in points, sharing ordinary text's finite -768 through 768 range.
- Preserve native hundredth-point precision, signed values and explicit zero independently of omission.
- Propagate through chart and trendline rich styles, precedence, source-bound edits and vector text without changing literal characters.
- Verify the complete field lifecycle with existing line/combo and vector fixtures.

## Capabilities

### New Capabilities
- `ppj-chart-text-letter-spacing`: Direct chart character spacing and its editing lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema, additive chart text style wire field, native style codec and PPJ mappings, focused regressions, generated manual/matrix and chart guidance. No new dependencies or host evaluation workflow.
