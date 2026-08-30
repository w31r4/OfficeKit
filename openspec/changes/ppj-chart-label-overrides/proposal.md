# Change: PPJ chart label overrides

## Why

PPJ owns chart-level labels, but one native series-specific `c:dLbls` or
point-specific `c:dLbl` currently makes the imported chart read-only. Agents
cannot emphasize one forecast, total, outlier, or threshold without replacing
the native chart with manual text.

## What changes

- Add one compact `series[].dataLabels` object with series defaults and sparse
  point overrides.
- Compile and project the bounded native `c:ser/c:dLbls/c:dLbl` profile.
- Permit capability-issued add, replace, and remove edits inside an otherwise
  editable ChartPart.
- Preserve unsupported custom label text, layout, shape/effect, leader-line,
  extension, and workbook-linked format graphs as source-owned.

## Impact

- Additive PPJ schema and Office wire messages; wire version remains 2.
- Shared NativeAOT ChartSpace reader/writer/projector/lowerer changes apply to
  both authored PPTX charts and bounded third-party chart continuation.
- One existing comprehensive PPJ contract gains authored/imported proof; no
  label matrix or new fixture suite is added.
