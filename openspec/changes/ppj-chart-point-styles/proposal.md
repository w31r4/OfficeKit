# Change: PPJ native chart point styles

## Why

PPJ can style a whole native chart series and can override one point's label,
but it cannot emphasize one measured bar, one pie slice or one doughnut segment
without rebuilding the chart as manual shapes. That is a common evidence-design
operation and a real gap between the language and native ChartML.

## What changes

- Add sparse `series[].pointStyles` keyed by zero-based logical point index.
- Support bounded fill, stroke and circular-slice explosion state.
- Compile and project canonical `c:ser/c:dPt` nodes.
- Permit capability-issued add, replace and remove edits while preserving
  series/data topology and all unrelated native content.
- Keep marker-specific, 3D, picture, extension and irregular point graphs
  source-owned and fail closed.

## Impact

- Additive PPJ schema and Office wire state; wire version remains 2.
- Shared NativeAOT ChartSpace code changes apply to authored PPTX and bounded
  third-party chart continuation.
- The existing comprehensive PPJ contract gains one point-style path; no
  effect matrix, new fixture farm or full test suite is added.
