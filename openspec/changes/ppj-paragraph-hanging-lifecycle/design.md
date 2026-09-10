## Context

See proposal.md. Authored hanging already negates the nearest-even EMU value; projection reverses the sign. Native deletion exists. The remaining gaps are PPJ authority/mask/lowering and native reading that currently trusts SDK integer coercion.

## Goals / Non-Goals

Complete ordinary text/shape hanging indent and preserve the left-indent lifecycle. Table/placeholder inheritance and host layout retain separate boundaries.

## Decisions

Reuse `hanging` and `no_indent`, retaining current sign and style-priority rules instead of introducing an alias. Add an exact field mask/authority and lower only changed paragraphs.

Parse raw native indentation and enforce the signed native limits. Remove the whole-paragraph layout gate now that both coordinates reject replacement locally, allowing unrelated edits to retain unmodeled values. Preserve unchanged coordinate spelling.

Generalize the left-indent fixture for both coordinates, retaining its assertions and adding signed hanging values. Reuse native XML/ZIP checks and malformed-source cases. Move the generic unsupported fixture to direct run styling, which retains its separate source-edit boundary.

## Risks / Trade-offs

Sign or zero is lost → assert both native integers and fresh PPJ, including signed fractions, endpoints and deletion/restoration.

Left indent regresses → run both coordinate fixtures and existing paragraph/table regressions.

SDK coercion normalizes malformed source → validate raw integers and verify no-op bytes, unrelated XML preservation and rejected replacement.

Preview is confused with host proof → retain partial diagnostics and separate NativeAOT/host evidence.
