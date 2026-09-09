## Why

F-07 chart text still rejects native cap attributes. Ordinary PPJ text already exposes capitalization, so an otherwise supported chart style cannot currently preserve the same display state without becoming source-owned.

## What Changes

- Add chartTextStyle.capitalization using the ordinary text none/small/all vocabulary.
- Retain explicit none separately from omission, and keep the literal text unchanged for every mode.
- Propagate through native chart styles, trendline rich paragraph/run/end styles, field precedence and vector text.
- Verify creation/change/reset/removal/recreation, invalid/unknown state and unrelated JS wire preservation with existing minimal fixtures.

## Capabilities

### New Capabilities
- `ppj-chart-text-capitalization`: Direct chart capitalization state and its edit lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema, additive chart style protobuf field, shared native style codec, authored/source-bound builders/projector, vector text, focused tests and chart documentation. No new dependency or host evaluation workflow.
