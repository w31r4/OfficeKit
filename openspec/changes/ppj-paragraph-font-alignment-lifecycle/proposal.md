## Why

F-03 still lacks direct paragraph font alignment. PPJ currently cannot express
how different font sizes align within a line, although DrawingML records this
independently of horizontal paragraph alignment and text-box anchoring.

## What Changes

- Add `text.paragraphs[].style.fontAlignment`: `auto`, `top`, `center`, `baseline`, `bottom`.
- Preserve explicit auto versus absence through authoring, native import and projection.
- Support ordinary text/shape source-bound add/set/remove/restore with exact field authority, preserving unknown native values during unrelated edits and refusing replacement.
- Update Help, generated references, registry, focused Agent guidance and backlog; prove the lifecycle with small native experiments and explicit preview limitations.

## Capabilities

### New Capabilities

- `ppj-paragraph-font-alignment-lifecycle`: direct paragraph font alignment and source-preserving editing.

### Modified Capabilities

None.

## Impact

PPJ schema, paragraph wire binding, shared native paragraph codec, authored and
source-bound compilation/projection, preview diagnostics, tests and presentation
documentation. Existing field numbers stay unchanged. Font metrics and host line
layout remain separate evidence; this change delivers codec source and generated bindings.
