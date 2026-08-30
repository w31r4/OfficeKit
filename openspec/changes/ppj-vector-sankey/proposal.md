# PPJ vector sankey

## Why

PPJ now owns every bounded chart family named by the local PPTD manual except
Sankey. The reference binary rasterizes Sankey to PNG, while OfficeKit already
has editable custom geometry, deterministic z-order and exact program recovery.

## What changes

- Add one finite `sankey` chart profile with explicit nodes and directed flows.
- Prove a directed acyclic, flow-conserving graph before output.
- Lower flows and nodes to an editable DrawingML group rather than an image.
- Document the semantic choice, native fallback and animation boundary.

## What does not change

- Arbitrary graph layout, cycles, negative flow and expression-driven styling
  remain unsupported.
- Imported shape ribbons are not guessed to be a semantic Sankey.
- Office wire version and ordinary ChartPart behavior do not change.
