## Context

PowerPoint commonly stores an SVG as a native `asvg:svgBlip` relationship
paired with a raster `a:blip` fallback. `PptxPictureCodec` already reads both,
requires the SVG MIME type, preserves the fallback, and replaces only the SVG
relationship after re-proving the fixed source topology. PPJ drops the second
asset during projection.

## Goals / Non-Goals

**Goals:**

- Make a paired SVG visible as a local, hash-bound PPJ asset.
- Preserve the raster fallback as ordinary `image.asset`.
- Replace only the SVG member through a capability-bound field edit.
- Recover the new vector asset after second projection.

**Non-Goals:**

- Embedding raw SVG XML inside PPJ.
- Adding an SVG node-edit DSL or XPath-like selectors.
- Generating a fallback raster from SVG.
- Adding, removing, or converting the native fallback-pair topology.
- Treating an ordinary standalone SVG image as a fallback pair.

## Decisions

### 1. One optional field represents the native pair

`image.svgAsset` names the exact `image/svg+xml` asset carried by an imported
fallback pair. `image.asset` continues to name the raster fallback. Authored
PPJ omits `svgAsset`; a source-free Agent can still use an ordinary SVG through
`image.asset` when that authored profile is supported.

### 2. Replacement is state, not an operation list

The source page issues `replaceSvg` for `image.svgAsset`. The Agent adds a new
local SVG asset declaration, points `svgAsset` to it, and keeps the nativeRef
unchanged. The compiler validates the bytes, MIME, hash, source revision, and
capability before calling the existing picture writer.

### 3. The fallback remains deliberately stale-looking but valid

Office uses the SVG when supported and retains the raster member for older
consumers. This slice does not silently rasterize a replacement SVG. Review
must inspect a native render and, when fallback behavior matters, test the
target legacy host separately.

## Risks / Trade-offs

- [Agent edits the raster instead] -> Keep `asset` and `svgAsset` distinct and
  document their native roles.
- [MIME or hash mismatch] -> Reuse the PPJ asset catalog validation and require
  `image/svg+xml` for `svgAsset`.
- [Topology drift] -> Issue capability only when import found a safe paired
  relationship; addition/removal stays unsupported.
- [False visual proof] -> Re-import and render, while recording that legacy
  fallback playback was not regenerated.

## Migration Plan

Additive optional image field plus one closed native capability operation and
field. Existing authored and source-bound PPJ remains schema-valid.

## Open Questions

None.
