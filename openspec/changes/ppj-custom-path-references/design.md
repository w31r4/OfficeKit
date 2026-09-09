## Context

See proposal.md. Native points/arcs already have reference slots; the PPJ helper currently reads only doubles and the projection gate requires literals. Shapes and image masks share the projection helper; standalone lines share the current pathCommand schema.

## Goals / Non-Goals

Carry every supported custom-shape path reference without evaluation/flattening. Keep heterogeneous/default path viewports and referenced masks/standalone lines as separate increments.

## Decisions

Use a customPathCommand schema variant so standalone line numeric behavior remains unchanged. Custom point numeric values keep viewBox-origin subtraction and 1000 native path units per unit; radius numbers use 1000 and angle numbers 60000 units/degree. Strings copy verbatim into native reference slots without viewBox-origin subtraction or unit conversion. Built-ins/adjustments/guides keep native evaluation units. Extend only shape projection/source gates; shared mask paths retain literal-only checks and lowering rejects references for non-shape owners. Existing source paths authority suffices; final native graph validation handles dangling and invalid arc values.

## Risks / Trade-offs

Mixed-unit confusion -> document native reference units and verify actual evaluation. Shared mask gate widened -> explicit shape-only flag and negative regression. Silent preview numeric coercion -> field diagnostics. Unsupported source viewport -> retain positive common viewport gate.

## Migration Plan

Additive custom path union; no wire changes. Fresh projections recover referenced paths. Existing numeric paths and lines keep their behavior.
