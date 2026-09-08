# Design

## Boundary

This slice adds one authored playback-start field to the existing media
primitive. `onClick` is the compatibility default. `onSlideStart` selects an
immediate `p:stCondLst/p:cond delay="0"` on the media's canonical
`p:cMediaNode`.

The field controls only the initial start condition. The existing media
click sentinel, payload, poster, trim, loop, mute, and package relationships
remain codec-owned. Imported media timing is not projected into this field,
because its timing tree may contain conditions or relationships that this
writer does not own.

## Lowering and recovery

The PPJ schema and `PresentationMedia.playback_trigger` wire field carry the
two-value contract. `PptxMediaCodec` validates it and `PptxTimingCodec` emits
the corresponding start condition. The embedded PPJ snapshot is the recovery
owner for authored intent; ordinary third-party PPTX import continues to
classify media as opaque/source-bound.

## Non-goals

- Editing imported media timing graphs.
- Parallel/sequence timing containers, conditions, media bookmarks, captions,
  volume/fade, or host playback verification.
- Payload or poster replacement.
