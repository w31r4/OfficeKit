## Why

OfficeKit already supports a bounded source-preserving continuation primitive:
an imported slide may retain its complete native object prefix and append
canonical textboxes, basic shapes, or an embedded rectangular image above it.
PPJ currently rejects every increase in a source-bound page's `elements` array,
so the public Presentation language hides this mature capability.

## What Changes

- Issue a page-level `appendElement/elements` capability for ordinary imported
  slides that can use the existing bounded-overlay writer.
- Let an Agent append typed `text`, bounded `shape`, and `image` elements to the
  end of a source-bound page's ordered `elements` array.
- Lower those new typed elements with the existing authored PPJ element
  compiler and the existing source-preserving PPTX writer.
- Reject interleaving, source identity loss, unsupported element kinds, and any
  overlay combined with another mutation in the same source slide.
- Teach Agents that appended elements are topmost overlays, not arbitrary
  insertion into the unknown native z-order.

## Capabilities

### New Capabilities

- `ppj-source-overlay-parity`: Capability-issued typed overlays on imported
  slides.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ capability vocabulary, projection, source-bound lowering, generated Skill
guidance, coverage, and one focused existing C# test are affected. The native
overlay writer, asset catalog, Office wire, JavaScript Presentation runtime,
and source graph preservation logic are reused without a second writer.
