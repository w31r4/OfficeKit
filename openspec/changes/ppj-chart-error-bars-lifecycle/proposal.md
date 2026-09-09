## Why

F-07 still lists source-bound error-bar additions and deletions as missing. The existing series.errorBars object supports four value types, but import/edit currently requires that the native object already exists.

## What Changes

- Support adding, changing, deleting and recreating the existing errorBars object on native bar/column and line series, including categorical combo series.
- Preserve trendlines, other series, chart data and unrelated ZIP entries.
- Validate existing native owners before deletion, and prevent an inserted scalar object from replacing unprojected custom plus/minus data.

## Capabilities

### New Capabilities
- `ppj-chart-error-bars-lifecycle`: Full presence lifecycle for the existing PPJ errorBars object.

### Modified Capabilities
None.

## Impact

PPJ source-bound compiler, PPTX topology checks, shared native error-bar patcher, focused tests and chart/backlog references. The existing schema, wire fields and XLSX adapter presence restriction stay unchanged. Custom error-bar data and formula/workbook synchronization remain separate gaps.
