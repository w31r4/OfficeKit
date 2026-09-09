## Why

F-07 in the PPJ gap backlog still treats a trendline label with nonempty manual layout as source-owned. Labels already expose text and style; their position and size need the same authored, source-bound and reprojection lifecycle.

## What Changes

- Add `data.series[].trendlines[].label.layout.manual` with native layout target, coordinate/size modes and numeric values.
- Preserve absent layout, empty layout and empty manual layout distinctly; reject unknown layout content without dropping source data.
- Add additive wire state and focused ordinary/combo chart lifecycle regressions; synchronize the chart reference and coverage entry.

## Capabilities

### New Capabilities
- `ppj-chart-trendline-label-layout`: Typed native trendline label layout and source-bound lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema, artifact protobuf/bindings, shared ChartML label codec, both PPJ compiler routes and projector, chart tests and presentation documentation. No new dependency or operation; use `setChartSeriesAnalytics`. Host PowerPoint positioning and full F-07 remain outside this increment's evidence.
