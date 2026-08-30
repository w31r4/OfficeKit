## Context

PPJ `image.mask` references the shared `geometry` schema. Preset geometry is fully wired to `PresentationImage.mask_preset`, while custom geometry is rejected in `PpjAuthoredPresentationCompiler`. The native codec already owns a bounded literal custom-path graph for shapes, including deterministic validation and DrawingML construction. PowerPoint stores a picture's visible silhouette in the same `p:pic/p:spPr` preset-or-custom geometry slot.

## Goals / Non-Goals

**Goals:**

- Reuse the existing PPJ path vocabulary and native custom-geometry compiler for image masks.
- Emit a real editable `a:custGeom` on `p:pic`, not a rasterized or SVG-clipped substitute.
- Reproject only the canonical literal path subset that maps exactly back to PPJ.
- Keep imported custom-mask topology read-only while allowing unrelated proven picture edits to preserve it.

**Non-Goals:**

- Arbitrary SVG paths, raw DrawingML formulas, guides, handles, connection sites, text rectangles, or 3D geometry.
- Source-bound custom-mask path editing.
- Boolean path operations or automatic vectorization of raster images.

## Decisions

1. **Add custom mask paths to `PresentationImage`, reusing `PresentationCustomGeometryPath`.** The path command and validation vocabulary stays single-sourced instead of creating a second mask language.
2. **Use a temporary `PresentationShape` adapter inside the picture codec.** `PptxCustomGeometryCodec` remains the only DrawingML custom-geometry reader/writer; the adapter transfers only canonical paths and rejects guides, handles, sites, and text rectangles.
3. **Allow exactly one geometry owner in picture shape properties.** Import accepts one supported preset or one supported custom geometry alongside the existing transform/border/shadow shell.
4. **Project only literal, common-viewBox custom masks.** The existing shape projection gate and JSON conversion are reused so ordinary import never invents a PPJ path from an irregular native graph.
5. **Keep source-bound topology immutable.** A custom mask can coexist with frame, crop, opacity, accessibility, or same-format image replacement, but any path difference fails closed.
6. **Extend one existing comprehensive test.** One authored custom mask proves native XML, import, PPJ projection, exact embedded recovery, and topology rejection; no path-command matrix is added.

## Risks / Trade-offs

- **[Picture codec regression]** Refactoring the preset-only geometry slot could weaken source checks. → Require exactly one preset-or-custom geometry and retain the existing child/attribute gates.
- **[Incomplete semantic projection]** Complex native custom geometry could be partially exposed. → Project only the literal no-guide/no-handle subset; everything else remains opaque/source-owned.
- **[Unexpected source mutation]** A diff could alter a custom path without a capability. → Compare the full ordered wire path graph and reject any source-bound topology change.
- **[Visual mismatch]** A valid path can still be aesthetically poor. → Keep this change at the file-semantic layer and require rendered review through the existing workflow.
