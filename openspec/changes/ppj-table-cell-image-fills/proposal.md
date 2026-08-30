## Why

PPJ already permits image fills anywhere the shared `fill` union is used, but authored table cells reject that valid state even though DrawingML tables and OfficeKit's native image-paint compiler can represent it. This leaves a visible schema/compiler gap and prevents image-led comparison grids, contact sheets, and evidence tables from being authored through the public PPJ language.

## What Changes

- Compile explicit and default table-cell image fills into native DrawingML `a:blipFill` while preserving the declared asset, fit, crop, tile, and opacity semantics.
- Carry table-cell image paint through the additive Office wire contract and validate it against the slide asset catalog.
- Keep imported unsupported table topology source-owned; this change does not broaden arbitrary source-bound table restyling.
- Remove image-filled authored table cells from the documented fail-closed boundary and teach Agents when a cell image is an information carrier rather than decoration.

## Capabilities

### New Capabilities

- `ppj-authored-table-image-fills`: Author, compile, validate, and recover PPJ table cells whose shared fill is an image.

### Modified Capabilities

None.

## Impact

- Additive protobuf v2 field on `PresentationTableCellFill`.
- PPJ authored compiler asset lowering and native PPTX table writer/validation.
- Existing comprehensive PPJ authored round-trip contract.
- Presentation capability registry, generated `ppj.md`, charts/tables guidance, and coverage evidence.
