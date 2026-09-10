## Why

F-03 direct paragraph default underline can be authored and projected, but cannot yet be independently edited through structured source-bound PPJ. Complete its explicit cancellation, deletion and restoration while preserving original text and native decoration graphs.

## What Changes

- Add independent defaultText.underline field authority for ordinary text/shape paragraphs.
- Reuse every existing underline token and single/double aliases; native XML uses sng/dbl and PPJ projection uses single/double.
- Support assignment, explicit none, deletion, underline-only wrapper removal and restoration.
- Preserve unknown source tokens and underline effect children during no-op/unrelated scalar edits; reject replacement even if effects have no explicit underline attribute.
- Extend focused lifecycle fixtures and synchronize public field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-underline-lifecycle`: Direct paragraph default underline presence and source preservation.

### Modified Capabilities

None.

## Impact

PPJ capability schema/validator/projector, structured paragraph mutations, default-run writer, tests, Help and references. Existing optional Underline suffices; no wire change. Underline effect graphs, placeholder inheritance and host glyph rendering retain separate boundaries.
