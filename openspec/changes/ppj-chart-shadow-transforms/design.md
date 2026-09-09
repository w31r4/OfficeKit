## Context

See proposal.md. `PresentationShadow` has eight fields. Shared chart builders and projector already preserve optional geometry, color/theme and alpha. The direct native shadow parser currently permits only blurRad/dist/dir/algn/rotWithShape. Its ordinary callers also prove siblings for other effect profiles.

## Goals / Non-Goals

Goals: carry the four missing direct outer-shadow transform attributes through the existing chart style lifecycle.

Non-goals: ordinary imported transform editing, new color graphs, general effect DAGs, preview rendering and host acceptance.

## Decisions

- Append optional sint32 fields 9–12 for scale ratios and skew angles. Reuse reflection transform units and bounds; absence carries native defaults.
- Give `TryReadOuterShadow` an explicit chart transform opt-in. Only `TryReadTextEffects`, used by the chart effect parser, enables it. Other owner proofs keep rejecting transform attributes.
- Shared native writer emits present attributes; chart compiler/projector propagate them. Existing vector clone/override paths must be verified rather than duplicated.
- Coexisting glow uses a write-only transform opt-in: chart reconstruction enables it, and run/default-run writers enable it only after a requested shadow has been rebuilt following the strict original-owner check. Ordinary glow reading and unrequested shadow mutation remain strict. This is necessary for vector text with shadow and glow together.
- Extend existing line/combo lifecycle, bounds/precision, named-style/vector and wire tests. Check coexistence with all four sibling effects in the same lifecycle.

The [Microsoft property contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outershadow.horizontalratio?view=openxml-3.0.1) defines signed scale/flip semantics. [Open XML SDK validators](https://github.com/dotnet/Open-XML-SDK/blob/main/generated/DocumentFormat.OpenXml/DocumentFormat.OpenXml.Generator/DocumentFormat.OpenXml.Generator.OpenXmlGenerator/schemas_openxmlformats_org_drawingml_2006_main.g.cs) use Int32 scales and exclude ±5400000 for skew. PPJ uses ratios and degrees to match existing reflection syntax.

## Risks / Trade-offs

- Shared parser affects ordinary sibling proofs → keep opt-in false and test ordinary no-transform guards.
- Tiny skew overflows after rounding → validate the native result; do not clamp.
- Concurrent preview work → isolate implementation and publish exact owned blobs.
- Structural success is not host appearance → retain F-07 and rendering residuals.
