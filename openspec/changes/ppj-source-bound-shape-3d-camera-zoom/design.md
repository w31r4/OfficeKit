# Design

- Owner: `p:sp/p:spPr/a:scene3d/a:camera/@zoom`.
- Native leaf: `shape3dSceneCameraZoomThousandthPercent`.
- The owner must be the single direct `a:scene3d` child of `a:spPr`. The
  scene must have exactly two direct children in schema order: a bare
  `a:camera` with recognized `prst` and one explicit canonical non-negative
  `zoom` token, followed by a bare `a:lightRig` with recognized `rig` and
  `dir` attributes. Neither child may contain a child element. The explicit
  zoom is expressed in DrawingML thousandths of one percent and is bounded to
  the native signed 32-bit range; an omitted zoom does not issue a leaf.
- Projection stores the zoom override in the native shape model and keeps the
  existing camera/light-rig leaves independently available. The edit plan
  rechecks the full owner topology and token-splices only `camera/@zoom`,
  preserving camera preset, light-rig state, root 3-D state, and all other
  package bytes.
- Missing, malformed, duplicate, extension-bearing, transformed,
  backdrop-bearing, unknown, default-only, or otherwise ambiguous scene state
  remains source-owned.

No source-free authoring, camera FOV/rotation editing, light-rig editing,
backdrop editing, relationship change, or full 3-D scene reconstruction is
included.
