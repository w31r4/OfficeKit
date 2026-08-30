## Why

PPJ imports bounded legacy comments, and the PPTX codec already edits their text
while preserving author, date, position, native identity, order, and package
topology. The projector currently omits comment nativeRef evidence and the
source-bound compiler treats the whole comments array as immutable.

## What Changes

- Bind each imported legacy comment to a stable nativeRef.
- Issue `replaceText` only when the owning slide's legacy comment profile is
  source-editable.
- Lower fixed-count, fixed-order comment text changes to the existing codec.
- Keep comment creation, deletion, reorder, metadata, modern threads, replies,
  anchors, and resolution outside this narrow slice.
- Extend an existing chart-edit transaction rather than add a new test path.

## Capabilities

### New Capabilities

- `ppj-legacy-comment-text-parity`: Source-bound PPJ editing of existing bounded
  legacy comment text.

### Modified Capabilities

None. No schema ID, protobuf, or wire version changes.

## Impact

Comment parsing/projection, source-bound root lowering, generated guidance,
coverage, and one existing PPJ contract are affected.
