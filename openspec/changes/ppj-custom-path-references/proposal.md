## Why

F-04 can now retain adjustment/guide/handle/site graphs but paths cannot reference them in PPJ. Expose native references in every custom-shape point and arc slot instead of making the whole shape opaque or flattening references.

## What Changes

- Custom shape move/line/quadratic/cubic coordinates and arc radii/angles accept numeric values or native reference strings.
- Preserve references across authored export, fresh projection and source paths edits; validate dependencies with the native graph.
- Keep positive shared viewports, line/mask/clip literal contracts and existing topology protections.

## Capabilities

### New Capabilities
- `ppj-custom-path-references`: Native reference identity in PPJ custom-shape path commands.

### Modified Capabilities
None.

## Impact

Custom path schema, authored lowering, projection/source geometry gate, focused native regression and preview diagnostics, Help/registry/Skill documentation. Native wire already supports references.
