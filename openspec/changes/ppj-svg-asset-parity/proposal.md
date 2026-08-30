## Why

OfficeKit already imports and source-preservingly replaces the SVG member of a
native PowerPoint fallback-picture pair. PPJ currently projects only the
raster fallback as `image.asset`, so an Agent cannot discover or change the
vector source even though the mature codec can prove that operation. This is a
high-value gap for SVG-led third-party decks and template continuation.

## What Changes

- Add optional source-bound `image.svgAsset` state.
- Project both the raster fallback and exact paired SVG as content-addressed
  PPJ assets.
- Issue a bounded `replaceSvg` capability only for a proven native SVG pair.
- Lower an `svgAsset` change through the existing residual-preserving picture
  writer while retaining the raster fallback, relationship topology, crop,
  geometry, effects, frame, and non-target package content.
- Teach Agents to edit or regenerate the local SVG asset, declare its new hash,
  and re-import the built PPTX.

## Capabilities

### New Capabilities

- `ppj-svg-asset-parity`: Truthful paired-SVG discovery and replacement.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

PPJ schema/model validation, projection, source-bound image lowering,
generated Skill guidance, coverage, and one existing comprehensive test are
affected. No new SVG syntax, wire field, or picture writer is introduced.
