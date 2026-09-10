# Add direct text reflection fade-angle leaf

## Why

PPJ reflection already carries `fadeAngle`, and authored text writes the native
`fadeDir` attribute. An imported direct rich-text run still treats every
reflection transform as source-owned, so an explicit full-span fade direction
cannot be edited independently.

## What Changes

- Expose `run.style.reflection.fadeAngle` as the
  `textReflectionFadeAngleDegrees` native leaf.
- Accept only a full-span direct run reflection whose sole modeled transform is
  an explicit canonical `fadeDir` token, and splice that token in the owning
  SlidePart.
- Keep scale, skew, alignment, rotate-with-shape, variable endpoints, unknown
  attributes, and other effect topologies source-owned.
- Add a focused source-bound edit/re-projection regression and update PPJ
  capability evidence.

## Capabilities

### New Capabilities

- `ppj-source-bound-text-reflection-fade-angle`: source-bound projection and
  token editing for one direct rich-text reflection fade direction.

### Modified Capabilities

<!-- No existing main spec is present; this change adds a bounded capability. -->

## Impact

The direct rich-text reflection reader, authored compiler/projector, native leaf
projection, and edit-plan proof/patch path gain one numeric leaf. The existing
protobuf reflection message already represents the value; the PPJ schema now
declares it for the direct reflection shape. No wire-version or relationship
changes are needed.
