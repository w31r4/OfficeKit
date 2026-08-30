## Why

PPJ already defines `pages[].name` and `pages[].hidden`, while the native PPTX
codec already edits `p:cSld/@name` and the bounded inverse `p:sld/@show` state.
Imported PPJ projects both values but does not issue page capabilities and the
source-bound compiler rejects their mutation. Two mature primitives are thus
documented as PPJ state without being usable for third-party continuation.

## What Changes

- Add page-level `setName` and use the existing `setHidden` capability.
- Issue `setName` for every hash-bound source page and `setHidden` only for a
  canonical visibility state.
- Lower capable name and visibility changes to the existing source-bound codec.
- Reuse one existing comprehensive PPJ edit instead of adding a test file or
  another package build.
- Synchronize generated PPJ guidance and coverage.

## Capabilities

### New Capabilities

- `ppj-page-metadata-parity`: Source-bound editing of PPJ page names and
  ordinary slide-show visibility.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ schema/model parsing, projection, source-bound lowering, generated guidance,
coverage, and one existing test are affected. No protobuf or OOXML writer change
is required.
