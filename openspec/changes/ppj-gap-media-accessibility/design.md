# Design

## Boundary

This slice treats imported media as a picture-shaped opaque object with one
independent non-visual metadata owner. It does not make the media payload,
poster, playback action, timing tree, or relationship closure editable.

## Lowering and recovery

The existing `PresentationMedia.accessibility` wire message is already the
typed owner used by authored media; `PresentationOpaqueElement.accessibility`
adds the same message for imported opaque owners. Import reads the bounded title,
description, and Office decorative extension through the residual
`PptxNonVisualAccessibilityCodec` profile. The PPJ projector places the value
on the common element field and emits `setAccessibility`; source-bound
compilation applies it back to the media picture's `p:cNvPr`.

The writer uses the residual-preserving path because media `cNvPr` may contain
the native media click sentinel or other children. Only the three modeled
metadata leaves are changed. An ambiguous extension graph remains unchanged
and has no capability.

## Non-goals

- Replacing or editing media bytes, poster images, or relationships.
- Inferring or rewriting `p:timing`, triggers, bookmarks, captions, or effects.
- Claiming playback or PowerPoint Accessibility Checker equivalence.
- Adding a new PPJ schema field or changing the Office wire version.
