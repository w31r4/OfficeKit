## Why

F-07 still rejects a trendline label whose number format carries `sourceLinked=true` or omits that native attribute. PPJ needs to preserve and edit that format state without claiming workbook synchronization or host display behavior.

## What Changes

- Add `label.numberFormatSourceLinked`: true/false for an explicit native flag, null for an omitted native attribute; require `numberFormat` when the field is supplied.
- Preserve existing omitted-PPJ-field behavior (explicit native false), and canonicalize explicit false back to the existing PPJ representation.
- Support authored, source-bound and fresh-projection lifecycle on ordinary and categorical combo trendlines, including format removal/recreation.
- Keep format state distinct from actual rendering: Microsoft documents that Office ignores this flag on trendline labels.

## Capabilities

### New Capabilities
- `ppj-chart-trendline-label-linked-format`: Complete native number-format link state for trendline labels.

### Modified Capabilities

None.

## Impact

PPJ schema, additive artifact protocol, shared label codec, both PPJ compiler routes/projector, focused native/JS regressions and presentation field documentation. Existing analytics operation; no workbook writes or new runtime dependency.
