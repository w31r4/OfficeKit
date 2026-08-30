## Why

PPJ v1 declares a typed `media` element, but the NativeAOT authoring compiler rejects every source-free media node. This makes the generated language reference materially overstate the usable surface and forces an Agent to fall back to static poster images for audio or video.

## What Changes

- Add an additive wire model for authored presentation media.
- Compile bounded local audio/video assets, poster frames, trim offsets, loop, mute, frame, transform, accessibility, and native click playback into canonical editable PowerPoint media.
- Keep imported unknown media source-owned and opaque unless the existing source-bound capability proves an edit safe.
- Recover exact authored PPJ and media assets through the existing embedded-program contract.
- Teach the PPJ reference, Presentations Skill, Help ownership, review guidance, and capability registry when to use media and where playback evidence remains necessary.
- Replace the false “media is modeled” completion claim with one narrow real round-trip and explicit format limits.

## Capabilities

### New Capabilities

- `ppj-authored-media`: Bounded PPJ audio/video authoring, canonical PPTX media packaging and playback state, exact authored recovery, and Agent-facing media guidance.

### Modified Capabilities

None. The active PPJ change has no promoted base spec; this change closes an implementation contradiction without altering the PPJ schema identifier or Office wire version.

## Impact

- `proto/office_kit/artifact/v1/office_artifact.proto` receives an additive `PresentationMedia` payload.
- PPJ parsing, semantic validation, authored lowering, asset validation, PPTX writing, timing writing, embedded recovery evidence, generated PPJ documentation, capability ownership, Skill guidance, and the existing comprehensive PPJ test are affected.
- No JavaScript authoring API, Office wire-version bump, raw OOXML surface, network fetch, media transcoder, or third-party source reconstruction is introduced.
