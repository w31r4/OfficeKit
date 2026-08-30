## Context

`SpreadsheetChartSeriesArtifact` already owns literal `x_values` and
`bubble_sizes`, and the PPTX codec has real scatter/bubble build, read and patch
paths. PPJ currently parses only shared categories and Y values, so it cannot
reach that native surface.

## Goals / Non-Goals

**Goals:**

- Make source-free scatter and bubble PPJ compile into editable native charts.
- Preserve numeric channels through PPTX reimport and third-party projection.
- Reject inconsistent vector topology before native export.

**Non-Goals:**

- ECharts-style tabular encodings, data filters, automatic grouping or remote
  data sources.
- Candlestick, heatmap, radar, waterfall or graph-series authoring.
- Expanding the existing source-bound `setChartData` authority to numeric X or
  bubble-size caches.

## Decisions

### Add direct finite vectors to each series

`xValues` and `bubbleSizes` remain parallel to `values`. This matches the native
owner and keeps the language deterministic. A generic `encode` object was
rejected because it would require a second table/query engine and substantially
larger validation semantics.

### Keep chart-family rules explicit

Scatter requires `xValues`; bubble requires both vectors; category charts reject
both. Shared `categories` must be empty for numeric-X charts. The compiler does
not infer channels from labels or array position.

### Preserve source authority

Projection exposes recognized numeric caches so Agents can inspect them. The
current `setChartData` capability still owns only series names and Y values;
attempting to change `xValues` or `bubbleSizes` remains fail-closed.

## Risks / Trade-offs

- [Risk] Agents expect a high-level data-frame grammar like PPTD. → Keep this
  slice native and finite; a later planning helper can generate the direct
  vectors without changing PPJ semantics.
- [Risk] Zero or negative bubble sizes render inconsistently. → Require finite
  positive sizes in schema and semantic validation.
