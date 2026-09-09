## Why

F-03 projects paragraph default text styles, but ordinary source-bound text/shape owners reject a change to defaultText.bold. Explicit false and removal need independent meanings without rewriting direct run styles.

## What Changes

- Extend ordinary text/shape paragraph-style authority with the existing PPJ path text.paragraphs[].style.defaultText.bold.
- Support true, false, omission and restoration, including removing a bold-only defaultText/style wrapper.
- Preserve fixed paragraph/run topology, unrelated defaults and direct run properties.
- Add minimal authored/source-bound/fresh-projection and permission regression evidence.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-bold-lifecycle`: Direct paragraph default bold editing.

### Modified Capabilities

None.

## Impact

PPJ source compiler, capability projection/validation, existing default-run codec and presentation documentation. Existing protobuf optional bold and default-style deletion intent suffice. Placeholder inheritance and other defaultText edits retain their separate boundaries.
