## Why

F-03 has explicit tab-stop lists but cannot express the direct default tab size
used by a paragraph. Imported `defTabSz` remains outside the modeled state.

## What Changes

- Add `text.paragraphs[].style.defaultTabSize` in points, preserving explicit zero and absence and the native signed 32-bit EMU range.
- Preserve authored precedence and ordinary text/shape source add/set/remove/restore, including single-field style removal, under exact field authority.
- Keep custom tab stops, indentation and literal tab characters independent; preserve unknown native tokens and non-target content.
- Synchronize discovery surfaces and add minimal native experiments plus explicit preview diagnostics.

## Capabilities

### New Capabilities

- `ppj-paragraph-default-tab-size-lifecycle`: direct paragraph default tab size and source-preserving edits.

### Modified Capabilities

None.

## Impact

PPJ schema, optional native wire field, shared paragraph codec, authored/source
compiler and projector, Help/registry/Skill, F-03 backlog and focused tests.
Measured tab placement, inherited defaults and host layout remain separate.
