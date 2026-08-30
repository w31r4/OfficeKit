## Context

PPJ already has a strict `media` union, local asset declarations, embedded-program recovery, and a canonical PresentationML timing writer. The missing bridge is the typed wire payload and native writer. PowerPoint represents one embedded video as a picture-shaped poster plus two relationships to one `MediaDataPart`, a `p14:media` extension, and a media timing node. Audio uses the same bounded package pattern with an audio relationship and timing node.

The implementation must remain deterministic, NativeAOT-compatible, source-free only, and honest about playback evidence. Imported vendor timing remains opaque.

## Goals / Non-Goals

**Goals:**

- Compile common embedded MP4, MP3, M4A/MP4-audio, and WAV assets with an explicit image poster.
- Preserve ordered PPJ z-order, frame, transform, accessibility, trim offsets, loop, and mute.
- Emit canonical editable PowerPoint media relationships and timing without raw XML in PPJ.
- Recover the exact PPJ and original media/poster bytes through the existing authored snapshot.
- Make the primitive discoverable in the generated language reference and focused Skill guidance.

**Non-Goals:**

- Network media, transcoding, duration probing, autoplay, bookmarks, fades, narration, media controls, or arbitrary timing graphs.
- Reconstructing imported third-party media as authored state.
- Claiming Keynote or PowerPoint playback from structural package validation alone.

## Decisions

### 1. Add one typed wire payload without changing wire version

`PresentationElement.media` carries media kind, media asset ID, poster asset ID, frame/transform, accessibility, trim offsets, loop, and mute. The field is additive to protocol v2. PPJ remains the only public authoring language; the wire object is an internal compiler IR.

Alternative considered: lower media to an opaque XML fragment. Rejected because it would expose native syntax, evade validation, and defeat stable capability ownership.

### 2. Use one content-addressed media asset purpose

Native assets use `asset/presentation/media/<sha256>`. The catalog validates MIME and magic bytes before package mutation. Supported profiles are MP4 video, MPEG audio, MP4 audio, and PCM/compatible WAV. Individual media uses the existing codec input ceiling rather than the 16 MiB image/OLE ceiling, while the envelope-wide uncompressed budget remains authoritative.

Alternative considered: accept arbitrary `video/*` and `audio/*`. Rejected because the package writer cannot prove the bytes match a PowerPoint-playable embedded profile.

### 3. Require an explicit poster image

Every authored media element requires `posterAsset`, and that asset must resolve through the existing safe image catalog. This keeps static rendering, accessibility review, z-order, and fallback behavior deterministic. OfficeKit does not synthesize or extract frames inside NativeAOT.

### 4. Emit the canonical editable PowerPoint graph

The writer creates one shared `MediaDataPart`, adds a media relationship plus an audio/video reference relationship from the slide, writes a poster `p:pic` with `ppaction://media`, writes `p14:media`, and targets the picture ID from a canonical `p:audio` or `p:video` timing node. Trim offsets use `p14:trim`; mute uses `p:cMediaNode@mute`; loop uses `p:cTn@repeatCount="indefinite"`.

When ordinary object animations and media coexist, the timing writer emits both under the same root timing graph. Existing unknown imported timing is never replaced by this authored path.

### 5. Recovery remains snapshot-exact

The existing `/officeKit/program.ppj` and program asset relationships already retain arbitrary declared asset MIME and bytes. Media therefore needs no second recovery format. A recovered authored deck returns exact PPJ and assets; ordinary third-party import continues to expose media as opaque/nativeRef state.

## Risks / Trade-offs

- [A structurally valid media part may use an unsupported codec profile] → Validate container signatures, document the boundary, and label playback evidence separately.
- [Large media inflates request memory] → Bound one media asset by `max_input_bytes` and all assets by `max_uncompressed_bytes`; do not add streaming in this slice.
- [Timing merge regresses ordinary animations] → Extend the one canonical timing writer and protect it with one combined media-plus-animation round trip.
- [Poster and media mismatch visually] → Require an explicit poster and make visual/playback review an Agent responsibility.
- [Schema again outruns implementation] → Mark the change complete only after compiled package relationships, exact recovery, generated docs, and registry ownership all pass together.

## Migration Plan

Add schema constraints, wire payload, compiler, writer, and guidance in additive commits. Existing PPJ files without media are unchanged. Rollback removes the additive media payload and restores the previous explicit authored-media rejection.

## Open Questions

None. Additional codecs and playback controls require separate observed evidence.
