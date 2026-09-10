## Why

F-03 still lacks editable list styling. PPJ already authors/projects `bullet.fontFamily`, but source-bound paragraph mutation rejects it and the native follow-text font choice is lost from PPJ.

## What Changes

- Complete the optional bullet-font choice: `fontFamily` or `fontFollowText: true`, mutually exclusive, for character, numbered and picture markers.
- Support ordinary source-bound text/shape font assignment, choice switching, removal and restoration under exact field authority. Preserve marker identity, other styling, runs, neighbors and non-target XML/ZIP.
- Align family validation at 1–255 XML-compatible Unicode scalars with non-whitespace content. **BREAKING**: reject the schema's formerly advertised 256-character value, already rejected by native validation for ordinary text.
- Retain unknown/ambiguous native font choices without flattening them; reject their replacement. Keep preview inheritance and glyph-layout limits explicit.

## Capabilities

### New Capabilities

- `ppj-bullet-font-lifecycle`: direct font choice authoring, projection, source editing, validation and preservation.

### Modified Capabilities

None.

## Impact

PPJ schema, authored/source compiler, projector/authority, shared bullet style codec, focused tests, Help, registry, generated manual/matrix and the text Skill reference. Reuse existing native wire choices; no protocol version change or new font dependency.
