# Design

## Public program

```json
{
  "type": "chart",
  "chartType": "line",
  "data": {
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [{
      "id": "revenue",
      "name": "Revenue",
      "values": [18, null, 31, 39]
    }]
  }
}
```

`null` means no observation at that logical category. It never means zero and
the compiler does not impute a value.

## Native representation

`SpreadsheetChartSeriesArtifact.values` retains one finite placeholder per
logical point. `missing_value_indexes` is strictly increasing, unique, and
bounded by `values.count`. The ignored wire placeholder is canonical zero.

The ChartSpace writer keeps `c:ptCount` equal to the logical count and omits
`c:pt` for missing indexes. The reader accepts only that bounded profile:
unique ordered indexes, finite present values, and a declared count no larger
than the existing point budget. Numeric X caches and bubble-size caches remain
dense.

## Imported continuation

Canonical sparse charts project to typed PPJ and retain their missing index
set in source semantics. A source-bound `setChartData` edit may change names
and present numeric values while the missing index set remains identical.
Changing present/missing topology is rejected because it would alter the
source cache graph rather than a proven scalar value.

## Verification

Extend the existing comprehensive PPJ build/project contract with one middle
missing point. Assert the authored cache shape, wire projection, PPJ `null`,
determinism, and source no-op. Do not add a chart matrix or a new fixture.
