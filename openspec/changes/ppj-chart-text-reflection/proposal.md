## Why

F-07 chart character styles cannot express the reflection already modeled for ordinary text. Add the same full-span reflection value with chart-native optional presence, so importing or independently changing an effect does not materialize missing properties.

## What Changes

- Add chartTextStyle.reflection with optional blur, startOpacity, endOpacity, distance and angle. An empty object retains the full-span effect; omission removes it.
- Carry it through shared chart typography, rich trendline styles, source-bound edits, named precedence and vector defaults/overrides, beside glow, inner shadow, outer shadow and soft edge.
- Preserve optional native value fields in the shared reflection codec; ordinary authored PPJ keeps its required fields. Unsupported spans/transforms and unknown effects remain source-owned.
- Correct the existing textReflectionDirectionDegrees leaf normalization so degrees use the declared native scale instead of the conflicting integer branch.
- Add focused lifecycle/presence regressions and update schema, wire, OpenSpec, capability references and backlog.

## Capabilities

### New Capabilities
- `ppj-chart-text-reflection`: independent full-span chart character reflection lifecycle.

### Modified Capabilities
None.

## Impact

PPJ schema, appended SpreadsheetChartTextStyleArtifact wire field, shared reflection codec, chart effects parser/writer, authored/source-bound compilation, projection, vector styles and focused tests/docs. This is an F-07 increment; rendering, transformed or partial-span reflection and host acceptance remain open.
