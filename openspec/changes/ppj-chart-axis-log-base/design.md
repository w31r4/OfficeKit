## Context

See proposal.md. Native charts share SpreadsheetChartAxisArtifact and XlsxChartAxisCodec across presentations and spreadsheets. PPJ already resolves numeric axis values through size grammar tokens.

## Goals / Non-Goals

Goals: preserve optional logarithmic scaling through authored and source-bound chart operations, including numeric x axes and combo secondary value axes.

Non-goals: change data caches, embedded workbooks, or chart topology when changing an axis scale.

## Decisions

- Add optional double log_base at wire field 25 and PPJ axis logBase. Absence represents linear scaling; a separate boolean would create conflicting states.
- Accept finite bases from 2 through 1000 and size grammar tokens. Explicit minimum/maximum must be positive for logarithmic axes. Category axes reject logBase.
  The [Open XML SDK ChartML schema](https://github.com/dotnet/Open-XML-SDK/blob/main/data/schemas/schemas_openxmlformats_org_drawingml_2006_chart.json) defines the 2–1000 inclusive constraint for CT_LogBase and places it before orientation in CT_Scaling.
- Write c:scaling/c:logBase before orientation. Patch only this scalar and preserve unrelated scaling children and package entries.
- Reject duplicate, malformed, or decorated logBase owners from editable native projection. Existing opaque preservation remains available.
- The shared codec also recognizes XLSX logarithmic axes. Preserve original logBase in the JavaScript spreadsheet wire adapter during unrelated edits, following its existing handling of native axis properties; this change adds no new spreadsheet authoring API.
- Carry the same field through radar spokeAxis.logBase. A first experiment showed that generic yAxis alone breaks the add/remove cycle when a linear radar axis reprojects as spokeAxis. Both authoring forms now share the same native value-axis field.

## Risks / Trade-offs

- Shared axis changes affect XLSX too → validate native read/write and PPJ source-bound paths with focused tests.
- Embedded PPJ could mask native omissions → strip embedded programs before every projection in the experiment.
- Passing package checks does not establish host rendering fidelity → report native XML and reprojection evidence only.
