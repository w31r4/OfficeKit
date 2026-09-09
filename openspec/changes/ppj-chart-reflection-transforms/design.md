## Context

See proposal.md. `PresentationReflection` has seven optional fields. `BuildChartTextReflection` already feeds authored, source-bound and vector chart styles; `ChartTextReflection` preserves presence. `PptxReflectionCodec` has a chart-only variable-position opt-in, while ordinary imported proofs call its default profile. Shared chart effect parsing already rejects duplicate/out-of-order lists.

## Goals / Non-Goals

Goals: complete the 14 direct DrawingML reflection attributes through the existing object and prove their independent lifecycle.

Non-goals: general effect DAGs, ordinary imported transform edits, preview painting or PowerPoint acceptance.

## Decisions

- Add wire fields 8–14 to the existing message: int64 fade angle, sint32 scale/skew pairs, optional string alignment and bool rotation. Presence is essential; materializing defaults would change imported XML.
- Keep scale as signed ratios and skew as signed angles. A percentage-only positive helper would lose mirrored scales. Use ties-even rounding; validate before conversion and after skew rounding, without clamping.
- Expand the chart opt-in to allow all known direct attributes. Default reader still rejects all transform attributes, even their native defaults. Reuse shadow alignment mapping with internal helpers.
- Extend the existing line/combo lifecycle, presence/bounds, style/vector and JS wire tests. Avoid a separate combinatorial acceptance suite.

Native property names and bounds were checked against [Microsoft Reflection documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.reflection?view=openxml-3.0.1) and the [Open XML SDK generated validators](https://github.com/dotnet/Open-XML-SDK/blob/main/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_drawingml_2006_main.g.cs). sx/sy are Int32; kx/ky exclude ±5400000; fadeDir is 0 inclusive to 21600000 exclusive.

## Risks / Trade-offs

- Shared wire/writer → ordinary imported negative tests retain the prior ownership boundary.
- Rounding near skew bounds → reject values rounding out of domain and test both signs.
- Concurrent preview edits → work in an isolated checkout and publish only exact owned blobs.
- Passing structural tests → report host display as unverified and leave F-07 open.
