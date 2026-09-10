## Why

F-03 has an authored/projected paragraph alignment field, but its source writer can overwrite unmodeled native alignment and its lowerer reapplies alignment to every paragraph. Complete the existing field with source-preserving assignment, deletion and restoration evidence.

## What Changes

- Retain `left`, `center`, `right`, `justify` and `distributed`; distinguish explicit left from omission.
- Patch alignment only on changed paragraphs under the existing exact field authority.
- Read native alignment tokens strictly, preserve unmodeled values during no-op and unrelated edits, and reject replacement.
- Prove text/shape lifecycle, precedence and surrounding XML/ZIP preservation with one focused fixture, then synchronize discoverability and backlog.

## Capabilities

### New Capabilities

- `ppj-paragraph-alignment-lifecycle`: direct paragraph alignment lifecycle and source preservation.

### Modified Capabilities

None.

## Impact

Source-bound PPJ lowering, the shared native paragraph-properties codec, focused tests, schema descriptions, Help, registry and generated references. Existing optional alignment already carries deletion through the complete paragraph state; no wire change is needed. Inherited placeholder/table editing and host text layout retain their existing boundaries.
