# Task: Images/figures placement + anchoring pitfalls

## Goal
Keep images and captions where you expect across Word/LibreOffice/PDF exports.

## Key reality
Image placement is the #1 LO-vs-Word mismatch.

## Inline vs floating
- **Inline** (`wp:inline`): behaves like a big character in the text flow. Most reliable for automation.
- **Floating/anchored** (`wp:anchor`): supports text wrapping, precise positioning, and "keep with paragraph" effects — also most likely to render differently between apps.

The public `DocumentModel.addImage(...)` API authors inline images by default.
Use `placement.type: "floating"` only for the bounded foreground anchor profile:

```js
document.addImage({
  dataUrl,
  alt: "Architecture overview",
  widthPx: 240,
  heightPx: 120,
  placement: {
    type: "floating",
    horizontal: { relativeTo: "margin", offsetPx: 260 },
    vertical: { relativeTo: "paragraph", offsetPx: 0 },
    wrap: "square",
    wrapSide: "bothSides",
    distanceFromTextPx: { top: 4, right: 12, bottom: 4, left: 12 },
  },
});
```

Horizontal references are `margin`, `page`, or `column`. Vertical references
are `margin`, `page`, or `paragraph`. Wrap is `square` or `topAndBottom`;
`wrapSide` is valid only for square wrap and may be `bothSides`, `left`,
`right`, or `largest`. Offsets are bounded to -10000 through 10000 pixels and
text distances to 0 through 10000 pixels.

## Recommendations
1. Prefer **inline** images for automation unless you truly need wrap-around.
2. For a floating figure that should move with nearby prose, prefer a
   paragraph-relative vertical position and keep the anchor image immediately
   before its caption/body paragraphs in logical document order.
3. Use `topAndBottom` when text must never sit beside the figure; use `square`
   only when side wrapping is intentional.
4. Use high-resolution sources and let Word scale down (avoid scaling up low-DPI images).
5. Keep a caption in a separate paragraph immediately after the image.

## Source-bound edit boundary

A recognized imported bounded anchor exposes `image.placement` through
inspect/resolve and permits only fixed-topology placement edits. It cannot be
changed from inline to floating or floating to inline after import. Behind-text
or overlapping anchors, tight/through wrapping, alignment/percentage position,
relative sizing, effect-bearing pictures, external images, and other irregular
drawing graphs stay source-owned and read-only. Do not rebuild them through the
semantic model; preserve them unchanged or fail closed.

The model SVG is an approximate planning preview. Word/LibreOffice native
rendering is authoritative for floating placement.

## Change one existing image description with an audit

When a reviewer has identified one recognized body image and supplied the
intended description, use the narrow public transaction rather than running a
document-wide metadata patch:

```bash
node examples/officekit-image-alt-text-edit-workflow.mjs \
  input.docx reviewed.docx image-alt.audit.json 1 \
  "Existing architecture overview" \
  "Architecture overview showing the API gateway and three service lanes"
```

The fourth argument is the inspected document **block index**. The workflow
requires the complete prior non-empty alternative text, so a stale locator or a
second matching-looking picture cannot silently receive the change. It accepts
only one canonical inline or bounded foreground floating PNG/JPEG image whose
`wp:docPr/@descr` and `pic:cNvPr/@descr` leaves agree. It keeps image bytes,
dimensions, placement, paragraph style, relationship graph, all other body
content, and every non-document part bound; only those two description leaves
inside `word/document.xml` may change. It also protects the input, refuses
output/audit overwrites, reimports, verifies/model-renders, and records a
byte-bound audit.

The operation deliberately does not add a missing description, decide whether
an image is decorative, alter the image title/name, or change placement/asset
bytes. Treat those as author-intent or package-review work. Render the source
and result through Word or LibreOffice plus Poppler and expect a zero-pixel
visual diff: alternative text improves nonvisual access, not page layout.

## Audit
```bash
python scripts/images_audit.py /mnt/data/input.docx
```

If you see `anchor` rows, treat as high-risk and inspect renders closely.

For a public-API result, also import the exported DOCX again and confirm the
image's `placement`, dimensions, alternative text, and surrounding logical
order before rendering.

## Render → PNG review checklist (images)
- Images appear on the intended page(s)
- No overlap with text, tables, or margins
- Captions remain adjacent to their figures
- Images aren’t blurry/pixelated (zoom to 200% to check)
- No unexpected cropping/stretching

## Common pitfalls
- Floating images shifting pages after small text edits
- Wrap modes causing overlap in LibreOffice exports
- Copy/pasted images with huge DPI metadata leading to surprising sizes
