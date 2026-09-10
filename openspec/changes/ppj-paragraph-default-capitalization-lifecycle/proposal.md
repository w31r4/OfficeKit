## Why

F-03 paragraph default capitalization is authored and projected but lacks independent structured source-bound editing. Complete its explicit cancellation and presence lifecycle alongside the other default scalar fields.

## What Changes

- Enable defaultText.capitalization assignment, explicit none, deletion, capitalization-only wrapper removal and restoration for ordinary text/shape paragraphs.
- Preserve the none/small/all enum, original text, direct run capitalization and unrelated default style/native content.
- Keep unmodeled native cap source-owned: no-op and unrelated scalar edits preserve it; replacement fails closed.
- Add focused regressions and synchronize public field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-capitalization-lifecycle`: Independent direct paragraph capitalization lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema/authority, paragraph diff/mutation, default-run writer, tests and references. Existing optional FontCaps suffices; no wire change. Host font rendering and inherited placeholders retain separate boundaries.
