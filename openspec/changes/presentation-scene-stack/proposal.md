## Why

OfficeKit already writes each slide as an ordered `PresentationElement` sequence, but the JavaScript model splits direct children into type-specific collections and reconstructs them in a fixed type order. That makes common image-led composition unreliable: an image added as a background is exported above previously authored text and overlays. It also hides the real source order of imported PPTX objects from Agents.

## What Changes

- Make the direct slide scene stack an ordered public collection shared by shapes, images, tables, charts, connectors, groups, and native objects.
- Add cross-type `sendToBack`, `bringToFront`, `moveBefore`, and `moveAfter` operations with stable inspection evidence.
- Add `slide.setBackgroundImage(...)` as a convenience for a full-slide, bottom-layer image while retaining the ordinary image object for editing and review.
- Preserve imported source order and expose source-bound reorder capability. Reorder only when the codec can independently prove the selected direct elements and their dependencies safe; otherwise fail closed without flattening unknown content.
- Teach Presentations and Presentation Template Creator to compose and review image-led layer stacks, contrast scrims, foreground readability, and obstruction risks.
- Reuse the existing three-file lossless benchmark and add layer-specific create/import/edit/export/reimport/render/continue evidence.

## Capabilities

### New Capabilities

- `presentation-scene-stack`: Ordered direct-slide elements, cross-type authored ordering, background-image composition, source-bound reorder capability, and layer inspection.
- `presentation-layer-acceptance`: Real-package preservation and visual evidence for authored and imported layer operations.

### Modified Capabilities

- Presentations Skill and Presentation Template Creator gain layer-aware composition and review guidance.

## Impact

- Changes the JavaScript Presentation object model, exporter, importer projection, Help, Skill guidance, and source-bound C# codec path.
- The existing ordered wire field is sufficient; Office wire remains version 2 unless implementation evidence proves otherwise.
- Does not expose raw OOXML, flatten opaque objects, or reinterpret a source-bound reorder as full imported-slide reserialization.
