# Authored media playback trigger

## Why

F-13 still leaves media playback triggers outside the PPJ language even
though the authored media writer already owns a canonical `p:cMediaNode`.
That forces authors to rely on an implicit click-only behavior and gives the
codec no typed field for the simplest automatic-play case.

## What Changes

- Add authored `media.playback.trigger` with the bounded values `onClick` and
  `onSlideStart`.
- Lower the field to the existing native media timing node without changing
  media payload, poster, or relationship ownership.
- Keep the default behavior compatible when `playback` is omitted.
- Leave imported/third-party media timing graphs source-owned and fail closed.
- Add a small authored XML and embedded-PPJ recovery experiment.

## Capabilities

### New Capabilities

- `ppj-authored-media-playback-trigger`: A bounded authored media start
  trigger for click-start and slide-start playback.

### Modified Capabilities

None.

## Impact

- PPJ v1 schema, generated artifact bindings, authored media parsing, and the
  existing timing writer.
- One focused presentation test plus media/F-13 coverage wording and Skill
  guidance.
- No Office wire-version change.
- Payload replacement, trim/volume/fade, captions, bookmarks, full timing
  graphs, and source-bound third-party playback edits remain outside this
  slice.
