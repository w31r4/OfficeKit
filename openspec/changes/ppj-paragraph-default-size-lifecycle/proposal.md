## Why

F-03 paragraph default size is authored and projected, but its structured source-bound assignment and removal are rejected. It needs an independent field lifecycle like default bold and italic.

## What Changes

- Authorize defaultText.size edits for ordinary text and shape paragraphs with fixed topology.
- Support assignment, omission, size-only wrapper removal and restoration, retaining other defaults and direct runs.
- Respect the native 1..768pt range and native hundredth-point rounding; reject out-of-range sizes.
- Add a focused lifecycle regression and update field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-size-lifecycle`: Source-bound paragraph default size presence and precision.

### Modified Capabilities

None.

## Impact

PPJ capability projection/schema/validation, source compiler, default-run writer and presentation documentation. Existing optional FontSizePoints suffices; no wire change.
