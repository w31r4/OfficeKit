## Why

F-03/F-06 spaceFirstLastParagraph can set booleans but cannot delete the direct source hint. Explicit false must remain distinct from absence.

## What Changes

- Add native deletion intent while preserving the existing optional boolean encoding.
- Complete PPJ removal/restoration and simple-style deletion on supported text owners and tables.
- Reuse boolean/wire experiments and publish field semantics and host-layout limits.

## Capabilities

### New Capabilities
- `ppj-text-first-last-spacing-removal`: first/last paragraph-spacing hint presence lifecycle.

### Modified Capabilities

None.

## Impact

Proto/generated JS, native validation/application/normalization, PPJ compiler, shared tests and public descriptions. Deletion requires updated codec.
