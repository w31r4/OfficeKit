## Why

F-03 already models numbered bullets, but `bullet.scheme` is an open string in the schema and ordinary imported paragraphs reject numbering-format edits. The 41 schemes supported by the native codec and the five `format` aliases need one discoverable, source-preserving PPJ contract.

## What Changes

- Enumerate the existing 41 native schemes and document the five alias mappings.
- Allow existing numbered text/shape paragraphs to change scheme or use a format alias, with exact authority for the requested spelling and canonical scheme reprojection.
- Preserve startAt presence/spelling, marker XML, styles, runs and neighbors; allow independently authorized scheme/startAt changes together.
- Retain required format selection, malformed-source preservation and explicit partial automatic-number preview diagnostics.

## Capabilities

### New Capabilities

- `ppj-bullet-numbering-scheme`: Closed numbering scheme catalog and source editing through canonical or alias syntax.

### Modified Capabilities

None.

## Impact

PPJ schema, source compiler/projector/semantic authority, shared bullet codec, focused tests, Help, registry, generated documentation, text reference and F-03 backlog. Existing wire and authored alias resolver are reused. Whole-marker replacement, numbering evaluation and inherited list layout retain their existing boundaries.
