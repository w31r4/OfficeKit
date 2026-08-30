## ADDED Requirements

### Requirement: Typed authored media
The NativeAOT compiler SHALL lower a source-free PPJ `media` element into typed presentation media state and SHALL preserve its stable ID, z-order, frame, transform, accessibility, media kind, local media asset, explicit poster asset, trim offsets, loop, and mute.

#### Scenario: Authored MP4 video
- **WHEN** a valid PPJ page declares an MP4 video and safe image poster with matching local asset hashes
- **THEN** build produces one editable native media picture at the declared z-order and records both assets in the receipt and embedded program map

### Requirement: Bounded media assets
The compiler SHALL accept only content-addressed local media whose declared MIME matches a supported MP4 video, MPEG audio, MP4 audio, or WAV signature, SHALL require a safe image poster, and SHALL reject missing, stale, mismatched, remote, oversized, or unsupported bytes before PPTX mutation.

#### Scenario: Media MIME disagrees with bytes
- **WHEN** a media asset declares `video/mp4` but lacks a valid ISO base-media `ftyp` signature
- **THEN** build fails with a path-specific asset diagnostic and produces no output PPTX

### Requirement: Canonical PowerPoint media package
The writer SHALL create one embedded `MediaDataPart`, the canonical media plus audio/video relationships, poster image relationship, `ppaction://media` picture state, `p14:media` extension, and a shape-targeted media timing node without exposing package relationship identities in PPJ.

#### Scenario: Video package relationships
- **WHEN** an authored video is compiled
- **THEN** the slide owns one video reference and one Office media reference to the same embedded media bytes, one poster image reference, and a timing target that resolves to the media picture ID

### Requirement: Bounded playback state
The writer SHALL express PPJ trim offsets, mute, and loop in the canonical media extension and timing node, SHALL bound trim values, and SHALL combine media playback nodes with existing authored object animations in one valid timing graph.

#### Scenario: Muted looping trimmed media with animation
- **WHEN** a page contains an animated text object and a media element with trim offsets, `loop: true`, and `mute: true`
- **THEN** the output retains both the object animation and media playback node, writes the trim offsets, mute state, and indefinite repeat, and reimports without an opaque-timing error

### Requirement: Exact authored recovery
The embedded-program contract SHALL recover the exact authored PPJ bytes and declared media/poster asset bytes from an OfficeKit-authored PPTX, while ordinary third-party media remains source-owned opaque content.

#### Scenario: Recover authored media program
- **WHEN** a built media deck is imported with its valid OfficeKit program parts intact
- **THEN** import restores the original PPJ, stable media element ID, media and poster files, and asset hashes without heuristic reconstruction

### Requirement: Honest Agent guidance
The generated PPJ reference, capability registry, and Presentations Skill SHALL describe supported media syntax, poster and format requirements, z-order behavior, review duties, and the difference between structural and real playback evidence.

#### Scenario: Fresh Agent needs a background video
- **WHEN** an Agent consults the PPJ media guidance
- **THEN** it can author a local video-plus-poster element, place overlays after it in the element array, and report playback as unverified until a real host playback check exists
