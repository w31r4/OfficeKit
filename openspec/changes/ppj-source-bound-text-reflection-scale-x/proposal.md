## Why

Imported rich-text runs still hide a direct reflection horizontal scale even though DrawingML stores it as a single scalar sx. That leaves a small, source-preserving edit gap in the F-03 text surface; a bounded field can close it without pretending to model arbitrary reflection graphs.

## What Changes

- Project a direct rich-text run reflection scaleX as the native leaf textReflectionScaleX when the reflection is full-span and contains only the canonical horizontal scale transform.
- Author and project the PPJ run.style.reflection.scaleX ratio using the existing 1/100000 native precision.
- Token-splice only a:reflection/@sx for a source-bound edit, preserving the run, sibling effects, and all non-target package parts.
- Keep scaleY, skew, alignment, rotateWithShape, fadeAngle combinations, variable endpoints, malformed tokens, and complex effect lists source-owned.

## Capabilities

### New Capabilities

- ppj-source-bound-text-reflection-scale-x: bounded authored and source-bound horizontal reflection scale for direct rich-text runs.

### Modified Capabilities

<!-- None. -->

## Impact

The PPJ schema and capability registry gain one ratio field and native leaf. The presentation authoring compiler, importer/projector, native leaf issuer, and edit-plan token splicer gain the bounded sx path. Coverage, presentation guidance, and the focused codec regression document the boundary. No protocol or package format version changes.
