# PPJ Frame Transforms

## Why

The PPJ `frame` contract already expresses rotation and horizontal or vertical
reflection, but the authored compiler rejects that state for charts, tables,
and groups. Imported graphic frames may carry the same native attributes while
projection currently drops them. This breaks template reconstruction and makes
the common frame vocabulary depend on element type.

## What Changes

- Add one reusable presence-aware frame-transform wire message.
- Preserve rotation and flips for native chart, table, and group frames.
- Compile and project PPJ frame transforms for those element types.
- Permit source-bound transform changes for safely modeled shapes, images,
  charts, tables, and groups under the existing `setFrame` capability.
- Keep connectors endpoint-driven and opaque objects fail closed.

## Impact

This is an additive wire-v2 change. It does not alter the strict PPJ schema or
introduce a second transform language. Unknown transform attributes and extreme
rotations remain source-owned.
