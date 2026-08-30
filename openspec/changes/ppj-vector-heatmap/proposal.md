## Why

PPJ can author eleven editable chart families, but analytical decks still need a
compact matrix carrier for intensity, risk, cohort, correlation and schedule
data. PPTD exposes a heatmap name by rasterizing the plot into one PNG and
embedding its source description. OfficeKit can retain the same declarative
convenience while producing a stronger native result: an editable DrawingML
group whose cells, labels and color scale remain ordinary PowerPoint objects.

## What Changes

- Add `chartType: "heatmap"` to PPJ with a bounded rectangular matrix profile:
  shared x categories, one named y-row per series, numeric or missing values,
  and an explicit heatmap style.
- Add linear and diverging color scales, optional domains and midpoint,
  editable value labels, cell gaps, missing-cell paint, borders and colorbar.
- Compile the semantic heatmap into one deterministic native DrawingML group,
  not a raster image and not a falsely labelled ChartPart.
- Preserve exact heatmap semantics through the embedded PPJ program; if that
  program is removed, ordinary PPTX projection truthfully exposes the native
  group instead of guessing that an arbitrary matrix of shapes is a heatmap.
- Teach Agents when a heatmap communicates a real matrix relation and when a
  table or ordinary chart is the honest carrier.

## Capabilities

### New Capabilities

- `ppj-vector-heatmap`: Author a bounded heatmap as editable native vector
  objects with deterministic scale and layout semantics.

### Modified Capabilities

None.

## Impact

- Additive PPJ v1 schema and C# semantic validation/compiler logic; no Office
  wire or protocol-version change.
- Existing comprehensive authored PPJ contract, generated `ppj.md`, focused
  charts/tables guidance, capability ledger and coverage evidence.
