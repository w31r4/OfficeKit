## Why

PPJ already describes image fills and image backgrounds, but the native authored compiler rejects shape image fills, tiled image elements, and any native image background beyond an opaque stretch. This makes the public language look smaller than PPTD in exactly the visual layer that matters for photographic masks, textured geometry, full-bleed backgrounds, and reusable design systems.

## What Changes

- Define one bounded image-paint profile shared by shape fills, slide backgrounds, and standalone images.
- Compile `stretch`, `cover`, `contain`, and deterministic default `tile` paint, explicit crop, and direct opacity into editable DrawingML rather than flattening the page.
- Project recognized imported shape/background image paint into PPJ with content-addressed assets and revision-bound capabilities.
- Permit only capability-issued source-bound edits to the bounded image-paint state; preserve and reject arbitrary blip effects, transforms, external links, and vendor topology.
- Keep z-order and semantic ownership explicit: a native background remains `p:bg`, a shape image fill remains shape paint, and an image element remains an independently ordered picture.
- Generate PPJ/Skill capability guidance and prove the profile in one authored/imported/edit/reimport sample.

## Capabilities

### New Capabilities

- `ppj-image-paint`: Bounded native DrawingML image paint for PPJ shapes, slide backgrounds, and standalone images, including projection and safe source-bound editing.

### Modified Capabilities

None.

## Impact

- Additive Office wire-v2 fields for reusable image-paint state on Presentation shapes and backgrounds.
- NativeAOT PPJ compiler/projector/differ and PPTX shape/background codecs.
- PPJ capability registry, generated `ppj.md`, focused Shapes and Media/Layers guidance.
- One existing integrated Presentation codec test and protobuf/generated evidence; no JavaScript Presentation authoring API is restored.
