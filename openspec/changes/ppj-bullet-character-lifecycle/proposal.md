## Why

F-03 exposes character bullets, but the schema permits up to eight characters while the native codec expects one Unicode scalar. Ordinary imported text/shape paragraphs cannot replace that scalar through PPJ, and shared native application rebuilds the marker instead of retaining its source XML.

## What Changes

- Align `bullet.character` with one XML-compatible Unicode scalar, including supplementary-plane symbols.
- Add exact source-field authority and replacement/restoration for existing character markers.
- Preserve marker attributes, bullet styling, paragraph state, runs, neighbors and non-target package content; preserve malformed/ambiguous source markers and reject their replacement.
- Document that the character is required for this marker kind and retain explicit preview limitations.

## Capabilities

### New Capabilities

- `ppj-bullet-character-lifecycle`: Single-scalar character marker validation and source-preserving edits.

### Modified Capabilities

None.

## Impact

PPJ schema, compiler/projector/semantic authority, shared bullet codec, lifecycle and preview fixtures, Help, registry, generated references, text guidance and F-03 backlog. Existing wire fields and preview painting suffice. Marker-kind changes, inherited bullet styles and host glyph/layout fidelity retain their existing boundaries.
