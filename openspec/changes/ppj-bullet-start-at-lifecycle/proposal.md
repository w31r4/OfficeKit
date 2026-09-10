## Why

F-03 exposes numbered bullets and their `startAt` value in authored PPJ, but ordinary imported text and shape paragraphs cannot edit that field through PPJ. A list starting at 3 must support changing, removing and restoring that direct value while retaining its original numbering scheme and surrounding source content.

## What Changes

- Complete `text.paragraphs[].style.bullet.startAt` for existing numbered paragraphs: integer 1–32767, explicit presence, deletion and restoration.
- Authorize that exact field through `setTextParagraphStyle`; retain list type, scheme, style and topology.
- Preserve unchanged lexical values and unknown XML; malformed or ambiguous source markers remain source-owned and reject replacement.
- Reuse the existing lifecycle fixture, update field guidance and retain explicit partial preview diagnostics for automatic numbering.

## Capabilities

### New Capabilities

- `ppj-bullet-start-at-lifecycle`: Direct auto-number start value lifecycle and source preservation.

### Modified Capabilities

None.

## Impact

PPJ schema descriptions and capability field vocabulary, semantic validation, source compiler/projector, shared bullet codec, focused native/preview tests, Help, generated references, capability registry and F-03 backlog. Existing wire fields suffice; no wire-version change or runtime dependency is needed. Character/picture markers, whole-bullet topology editing and inherited numbering evaluation retain their existing boundaries.
