## Why

F-07 chart text cannot yet express direct highlight paint, although ordinary PPJ text already accepts highlight colors. Imported highlight elements therefore make otherwise supported chart styles source-owned.

## What Changes

- Add chartTextStyle.highlight using the existing PPJ color and color-token vocabulary, resolving to opaque RGB as ordinary text does.
- Preserve highlight independently of text fill and typography, with creation/change/removal/recreation and canonical color projection.
- Propagate through chart and rich trendline styles, field precedence, source-bound edits and vector text.
- Verify native highlight order, unchanged characters and target-only edits in existing line/combo fixtures, plus one vector override fixture.

## Capabilities

### New Capabilities
- `ppj-chart-text-highlight`: Direct RGB chart text highlight and its editing lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema, additive chart text wire field, shared native style codec, authored/source-bound/projector/vector mappings, focused regressions and chart documentation. Alpha and unknown native color graphs retain existing fail-closed behavior; no host appearance claim or new dependency.
