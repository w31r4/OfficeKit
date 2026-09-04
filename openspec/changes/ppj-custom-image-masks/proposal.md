## Why

PPJ uses the same bounded geometry union for shapes and image masks, but authored images accept only preset masks even though OfficeKit already compiles the custom path vocabulary for native DrawingML shapes. This prevents an Agent from expressing irregular editorial crops, branded silhouettes, and template-derived image contours without rasterizing the composition.

## What Changes

- Compile PPJ `image.mask.kind: "custom"` through the existing bounded custom-geometry path vocabulary into native picture shape properties.
- Carry canonical custom mask paths through the additive Office wire, picture validation, import, and PPJ projection.
- Preserve custom-mask identity during unrelated source-bound edits, while
  allowing a capability-issued transition between a recognized
  no-adjustment preset and bounded literal custom geometry.
- Remove custom-path image masks from the authored fail-closed registry boundary and add focused Agent guidance.

## Capabilities

### New Capabilities

- `ppj-custom-image-masks`: Author, compile, recover, and source-preserve bounded custom path masks on native PowerPoint pictures.

### Modified Capabilities

None.

## Impact

- Additive protobuf v2 fields on `PresentationImage`.
- PPJ authored compiler, native picture codec, semantic projection, and source-bound guard.
- Existing comprehensive PPJ round-trip contract, capability registry, generated language manual, media/layers guidance, and coverage evidence.
