## Why

F-03 paragraph default language is authored and projected, but structured source-bound edits still fail. It needs direct presence semantics independent from run language and font defaults.

## What Changes

- Add exact defaultText.language authority for ordinary text/shape paragraph assignment, deletion and restoration, including language-only wrapper removal.
- Reuse bounded language-tag validation and preserve spelling, other defaults, direct run languages and unknown native attributes.
- Reject replacement of a source language outside the modeled tag grammar; preserve that native value during unrelated scalar edits.
- Extend the shared lifecycle fixture and public field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-language-lifecycle`: Direct paragraph default language presence.

### Modified Capabilities

None.

## Impact

Capability projection/schema/validation, source paragraph mutation, default-run writer, references and tests. Existing optional Language field suffices. Host proofing, full BCP-47 registry validation and language inference remain separate.
