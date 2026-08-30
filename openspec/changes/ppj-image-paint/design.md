## Context

PPJ has one `fill` union whose image branch already exposes asset, fit, crop, and opacity. The authored compiler currently lowers this branch only for an opaque stretch-only slide background, rejects it on shapes, and rejects `tile` on image elements. Imported shapes with `a:blipFill` therefore become opaque even when the native graph is simple and deterministic.

The native wire has separate legacy fields for pictures and slide backgrounds and no shape image-fill state. The new path must remain additive in wire v2, keep unknown blip graphs source-owned, and avoid reintroducing a JavaScript Presentation object model.

## Goals / Non-Goals

**Goals:**

- Make the existing PPJ image-fill vocabulary executable for source-free shapes, slide backgrounds, and picture elements.
- Preserve semantic ownership and editability in native PPTX: `p:bg`, `p:spPr/a:blipFill`, and `p:pic/p:blipFill` remain distinct owners.
- Project a bounded imported profile and permit revision-bound edits only when the source graph still matches.
- Keep authored and imported state discoverable through one generated PPJ/Skill description.

**Non-Goals:**

- Arbitrary blip effects, color transforms, artistic effects, external links, pattern fills, theme fill references, or vendor extensions.
- User-configurable tile scale, offset, alignment, mirror, DPI, or arbitrary fill rectangles in this tranche.
- Inferring whether an imported source rectangle originally came from `cover`, `contain`, or a manual crop.
- Flattening a shape or background to a full-page picture.

## Decisions

### Use one bounded native image-paint message

Add an additive `PresentationImagePaint` wire-v2 message with content-addressed asset ID, optional signed source rectangle, optional direct alpha, and a `stretch` or default `tile` mode. `PresentationShape` and `PresentationBackground` reference this message; `PresentationImage` adds only the same bounded mode because it already owns asset, crop, and alpha fields.

This is preferred to four unrelated scalar additions because validation, DrawingML ordering, relationship cleanup, and projection rules then have one owner. Existing background image fields remain readable for wire compatibility but new PPJ compilation emits the shared message.

### Lower high-level fit before the wire boundary

`stretch` writes no source rectangle. Explicit PPJ crop wins over fit. Otherwise `cover` and `contain` deterministically lower to a signed source rectangle using declared pixel dimensions and the destination frame. `tile` writes a parameter-free native tile node. Imported stretch paint projects as `fit: "stretch"` plus any exact crop; imported default tile projects as `fit: "tile"` plus any crop.

This accepts the information loss inherent in native DrawingML: a source rectangle does not prove whether the author intended cover, contain, or manual crop. Exact authored PPJ recovery still retains the original high-level choice through the embedded program.

### Keep direct alpha narrow

Opacity maps only to one `a:alphaModFix/@amt` child on the embedded blip. Any second blip effect, missing/extra attributes, unsupported compression/link state, or effect container makes the native paint unrecognized and source-owned.

### Capability-bind imported changes

Recognized shape image fills use the existing `setFill` capability. Recognized page backgrounds receive `setBackground`; recognized tiled picture elements receive `setImageFit`. The source-bound compiler must reproject the exact source, verify source/capability hashes, and change only the declared semantic owner. Relationships are added or removed through the existing part context and unreferenced assets are cleaned conservatively.

### One integrated contract sample

Extend the existing comprehensive PPJ codec test with an authored image-filled shape, a cropped translucent native background, and a tiled picture. Reimport it, edit one imported shape fill and one background, and verify the unaffected owners remain stable. Do not add an effect matrix or new test harness.

## Risks / Trade-offs

- [Negative source rectangles render differently in some non-Office hosts] → Retain existing signed-crop contract, render the integrated sample, and document host variance rather than rasterizing.
- [Changing fill relationships can orphan media parts] → Use `PptxPartContext` relationship accounting and remove only proven unreferenced relationships.
- [Default native tile sizing may not match every design expectation] → Expose deterministic default tile now; defer tile transforms until PPJ has explicit typed parameters.
- [Imported intent cannot be reconstructed from equivalent native crop] → Project exact executable state and preserve authored intent only through embedded PPJ recovery.
- [Broader shape editability could accidentally absorb unsupported effects] → A shape is typed only when its entire direct fill matches the bounded profile; otherwise it remains opaque or receives narrower native leaves.
