## Why

F-07 explicitly lists source-bound trendline additions and deletions as missing. PPJ already has the complete trendlines array syntax, but three count/topology guards prevent users from adding a fit or removing one after importing a native chart.

## What Changes

- Make data.series[].trendlines a writable ordered list on recognized native bar/column and line series, including those in categorical combo charts.
- Support insertion, reordering, replacement, removal of individual entries, and clearing the field with an empty array or omission.
- Preserve series identity, data/cache/workbook content, error bars and unrelated chart/package content.
- Continue rejecting unknown or decorated native trendline owners and invalid trendline parameters.

## Capabilities

### New Capabilities
- `ppj-chart-trendline-lifecycle`: Source-bound editing of the existing native trendline list through its full add/edit/remove lifecycle.

### Modified Capabilities
None.

## Impact

PPJ source-bound compiler, PPTX topology checks, shared ChartML trendline patcher, focused tests, capability registry and references. Existing schema and wire shape remain valid. XLSX retains its existing adapter-level list-count restriction; error-bar topology and trendline label/effect graphs are separate gaps.
