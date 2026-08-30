## Why

PPJ already stores sections and custom shows, and the native PPTX codec already
edits their names and page membership within a hash-bound fixed topology.
Imported PPJ currently projects only plain values, then rejects every change at
the root. Mature native behavior is therefore invisible to an Agent.

## What Changes

- Add `nativeRef` to PPJ section and custom-show state.
- Add a bounded `setPages` capability while reusing `setName`.
- Issue both capabilities only for native editable fixed-topology graphs.
- Lower capable edits to the existing section/custom-show codecs.
- Extend the existing comprehensive PPJ contract rather than add a test file.
- Synchronize generated PPJ guidance and coverage.

## Capabilities

### New Capabilities

- `ppj-section-custom-show-parity`: Source-bound name and page-membership edits
  for canonical PowerPoint sections and custom shows.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ schema/model parsing, projection, semantic validation, source-bound
lowering, generated guidance, coverage, and one existing test are affected. No
protobuf or OOXML writer change is required.
