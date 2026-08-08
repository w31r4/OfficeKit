# Post-edit review contract

Use this contract only after the final file has been exported and reopened.
Native OfficeKit inspection remains the authority for package, model, source,
and render facts. AnyDoc is an optional reading view.

## Review order

6. **Semantic review:** compare the user's requested facts and changes with the
   reopened model and the narrowest inspect evidence. `reviewArtifact()` checks
   modeled semantic integrity; the Agent must still judge whether the user's
   request was satisfied.
7. **Structural review:** inspect the actual DOCX/XLSX/PPTX/PDF bytes, package
   relationships, source bindings, opaque content, and unexpected parts.
8. **Layout/render review:** run deterministic layout and render checks for
   blank output, invalid geometry, clipping, overflow, overlap, crop, and
   baseline differences supported by the owning Skill.
9. **Optional content view:** request AnyDoc Markdown only when it materially
   improves review.
10. **Visual or human review:** understand the rendered pixels when that
    capability exists. Otherwise report the limitation honestly.
11. **Delivery review:** prove the final absolute path, type, SHA-256, distinct
    source/output paths, and evidence locations.

## Decide whether to use AnyDoc

AnyDoc overlaps with visual review only in confirming that some text or table
content exists. It does not preserve exact placement, typography, crop,
contrast, visual hierarchy, chart appearance, or slide composition.

Use `contentView: "anydoc"` when any of these is true:

- the Agent cannot directly understand rendered images;
- a long document, workbook, deck, or PDF would be expensive to read through
  page images;
- a compact cross-format Markdown view makes semantic comparison easier;
- an independent parser view is useful for a suspicious content mismatch.

Omit `contentView` for a small targeted edit when native inspect evidence and a
direct visual review already answer the task. Do not run AnyDoc merely because
it is installed.

AnyDoc is not OCR and is not a substitute for render review. A host-supplied
OCR result or image description is also derived text evidence, not direct
pixel review. If imagery, layout, or aesthetics are material and the Agent
cannot understand the render, use `visualReview: "requires-human"`; otherwise
use `"unavailable"`.

## Public API

```js
const { reviewArtifact } = await ctx.import("office-kit");

const review = await reviewArtifact(outputPath, {
  source: inputPath,
  contentView: "anydoc", // omit when the text view is unnecessary
  visualReview: "unavailable",
  maxContentChars: 40_000,
});
```

`reviewArtifact()` returns `semantic`, `structural`, `layout`, `contentView`,
`visualReview`, `delivery`, and one bounded Markdown `summary`. Its `verdict`
is `passed`, `passed-with-limitations`, or `failed` for the machine review. Read
the summary and compare it with the request before delivery.

AnyDoc loads only after `contentView: "anydoc"`. If that explicit view is
unavailable or unsupported, record the reason and continue the native review;
do not silently choose OCR, another parser, or another editing engine.
