## Why

PPJ already declares `chartType: "waterfall"`, but the authored compiler still rejects it. This is both a misleading language promise and a visible gap in analytical and management-report authoring where cumulative bridges are a common data carrier.

## What Changes

- Define bounded PPJ waterfall semantics with one numeric series and an ordered `pointRoles` vector of `delta` or `total`.
- Add explicit increase, decrease, and total role styles under the existing chart style object.
- Compile the semantic series deterministically into one native editable stacked-column chart with an invisible offset series and three visible role series.
- Preserve the exact authored PPJ through embedded-program recovery while leaving third-party ChartEx and irregular waterfall-like charts source-owned.
- Remove waterfall charts from the authored fail-closed registry boundary and teach Agents the supported analytical use and honesty constraints.

## Capabilities

### New Capabilities

- `ppj-waterfall-charts`: Author and recover a bounded cumulative bridge as a native editable PowerPoint chart without using shapes or raw OOXML.

### Modified Capabilities

None.

## Impact

- Additive PPJ v1 schema and typed C# model fields.
- PPJ semantic validation and authored chart lowering; no Office wire version change.
- Existing comprehensive PPJ chart contract, capability registry, generated `ppj.md`, focused chart guidance, and coverage evidence.
