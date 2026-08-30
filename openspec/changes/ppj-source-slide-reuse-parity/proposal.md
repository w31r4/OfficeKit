## Why

OfficeKit already proves and performs bounded source-preserving slide reuse:
it clones the SlidePart plus its closed owned graph, rebinds proven shared
resources, and keeps unknown native descendants intact. PPJ currently rejects
every source-bound page insertion, so an Agent cannot express this mature and
high-value continuation primitive in the public presentation language.

## What Changes

- Add a finite page-level `sourceClone` descriptor to PPJ.
- Issue a page `duplicate` capability only when the imported slide's native
  clone analysis is known and supported.
- Lower one unchanged clone placed immediately after its retained source page
  through the existing `PresentationSlide.clone_source` codec path.
- Require a build/reimport boundary before the cloned page can be edited.
- Preserve sections and custom shows unchanged during the pending clone build.
- Teach Agents that `sourceClone` is a bounded source macro, not a general
  page-copy algorithm or editable snapshot.

## Capabilities

### New Capabilities

- `ppj-source-slide-reuse-parity`: Capability-issued exact source slide reuse.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ schema/model validation, projection, source-bound page lowering, generated
Skill guidance, coverage, and one existing comprehensive test are affected.
The native clone codec, Office wire schema, JavaScript runtime, and authored
PPJ compiler are reused without a second cloning implementation.
