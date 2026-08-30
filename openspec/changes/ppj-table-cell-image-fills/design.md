## Context

The shared PPJ `fill` union includes image paint with an asset reference, fit mode, crop, tiling, and opacity. `tableCell.fill` and `tableStyle.defaultCellFill` both reuse that union, but `PpjAuthoredPresentationCompiler` currently lowers only none, solid, and gradient fills. Native DrawingML table-cell properties accept `a:blipFill`, and OfficeKit already has `PresentationImagePaint`, `PpjImagePaintLowering`, `PptxImagePaintCodec`, and a slide-scoped asset catalog.

## Goals / Non-Goals

**Goals:**

- Lower explicit and inherited PPJ table-cell image fills through the existing image-paint semantics.
- Emit a native editable `a:blipFill` owned by the slide part with deterministic asset relationships.
- Preserve PPJ exactly through the authored embedded-program recovery path.
- Keep the capability and Agent reference synchronized with the real compiler surface.

**Non-Goals:**

- Inferring or editing arbitrary imported table-cell image fills.
- Adding remote assets, raw relationship IDs, or raw DrawingML.
- Expanding table topology, multiple header rows, or other table-style boundaries.

## Decisions

1. **Reuse `PresentationImagePaint` on the table-cell fill wire union.** This keeps crop, tile, stretch, opacity, and asset validation identical to shape/background image paint. A second table-specific image model would duplicate semantics and drift.
2. **Resolve the image at PPJ lowering time using the cell's physical frame.** Each explicit or inherited cell fill is lowered after row and column extents are known, so `cover` and `contain` crops are computed against the actual cell dimensions.
3. **Build the relationship through the existing slide `PptxPartContext`.** Table-cell `a:blipFill` references media owned by the slide part and uses the same content-addressed asset catalog as other authored images.
4. **Keep source-bound table editing unchanged.** `PptxTableCodec.TryRead` intentionally accepts a bounded text/topology profile. Existing unsupported image-filled cells remain opaque/source-preserved instead of being reconstructed from incomplete semantics.
5. **Use one existing comprehensive PPJ round-trip test.** The contract needs one explicit image-filled cell plus one inherited default image fill, native XML/relationship evidence, and exact PPJ recovery; no effect/fit matrix is added.

## Risks / Trade-offs

- **[Relationship duplication]** Multiple cells may reference one asset. → Reuse the slide context's deterministic asset relationship resolution.
- **[Incorrect crop geometry]** Default fills can be inherited by cells with different dimensions. → Lower each cell independently using its computed row/column extents.
- **[Fallback projection remains opaque]** Removing embedded PPJ can lose authored fill semantics during ordinary projection. → Keep the current fail-closed imported-table boundary and state it explicitly; exact authored recovery remains available through embedded PPJ.
- **[Large image-heavy tables]** Many cells can create visually noisy or heavy decks. → Reuse the existing local asset once and add concise Skill guidance that images must carry evidence or identity.
