## Context

`OpenXmlChartTrendlineLabelCodec` currently permits only empty `c:layout`; the two PPJ compiler paths and projector already share its typed label state. The previous label tests cover ordinary line and categorical combo owners with exact non-target ZIP preservation. See proposal.md for motivation.

## Goals / Non-Goals

Implement the complete modeled manual-layout tuple for this label owner and its lifecycle. Other label text/effect graphs, other chart layout owners and host rendering remain separate gaps.

## Decisions

- Use `label.layout` with optional `manual` object, rather than a synthetic presence flag. `{}`, `{manual:{}}` and omission represent the three native container states directly.
- Manual fields are `target` (`inner`/`outer`), `xMode`, `yMode`, `widthMode`, `heightMode` (`edge`/`factor`), and finite `x`, `y`, `width`, `height`. Numeric values are native chart fractions, including zero, negatives and values outside 0–1; no ungrounded range clamp or slide-unit conversion. Optional values remain absent. A present native mode/target leaf without `val` projects its XSD default (`factor`/`outer`); physical no-op retains source bytes.
- Append two protobuf messages and add label field 6. Use optional strings/doubles for leaf presence and message presence for containers. Preserve established descriptor ordering.
- A small shared layout codec handles wire validation, ordered strict XML, and mechanical PPJ conversion, reused by both compiler routes and projector. The existing label semantics comparison and analytics edit operation carry lifecycle changes.
- Unknown attributes/children, extensions, duplicates, invalid order, non-finite numbers and unexpected text fail closed. Do not make malformed source editable by stripping unknown data.

Native semantics references: [Microsoft manual layout](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.manuallayout), [Office layout modes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/bb8565e7-dfea-453f-868b-ae680ab8a823), and [ChartML XSD defaults](https://github.com/plutext/docx4j/blob/master/xsd/dml/dml-chart.xsd).

## Risks / Trade-offs

- Width/height with `edge` denote an edge position, not a size → preserve modes and document the distinction; do not reuse slide frame helpers.
- Previously rejected empty manual layouts become editable → update old unsupported fixtures to contain real extensions and retain opaque/no-op regressions.
- XML schema and field tests do not prove host placement → report only native structure and round-trip evidence.
