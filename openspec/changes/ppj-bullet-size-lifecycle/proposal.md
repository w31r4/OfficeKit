## Why

F-03 list size still rejects ordinary source-bound PPJ edits, omits native follow-text size, and advertises wider numeric ranges than the native codec accepts. The three native size choices need a consistent PPJ lifecycle.

## What Changes

- Complete the mutually exclusive `bullet.size`, `sizePercent` and `sizeFollowText: true` choices for character, numbered and picture markers. Omission removes the direct size declaration.
- Align points to 1–768 in native 0.01-point units and relative size to 0.25–4 in native 0.00001 units. `sizePercent: 1` means 100%, distinct from follow-text and absence. **BREAKING**: schema-only values outside native ranges are rejected during validation.
- Support add/set/switch/remove/restore for ordinary source-bound text/shape markers under exact authority for every changed property.
- Preserve marker, font/color, runs, neighbors and non-target XML/ZIP. Retain malformed or ambiguous source sizes and reject replacement; invalid numeric tokens must not abort projection.
- Keep direct point-size preview and explicit relative/follow-text layout limitations discoverable.

## Capabilities

### New Capabilities

- `ppj-bullet-size-lifecycle`: direct size choice validation, source editing, projection and preservation.

### Modified Capabilities

None.

## Impact

PPJ schema, authored/source compilers, projector/authority, shared bullet style codec, focused native/preview tests, Help, registry/manual/matrix, text guidance and F-03 backlog. Reuse existing wire choices; no protocol version or dependency change.
