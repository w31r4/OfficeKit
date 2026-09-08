# Design

- Owner: `p:sp/p:spPr/a:scene3d/a:camera/@fov`.
- Native leaf: `shape3dSceneCameraFov60000`.
- The owner must be the single direct `a:scene3d` child of `a:spPr`. The
  scene must have exactly two direct children in schema order: a bare
  `a:camera` with recognized `prst` and one explicit canonical FOV token,
  optionally alongside one explicit canonical zoom token, followed by a bare
  `a:lightRig` with recognized `rig` and `dir` attributes. Neither child may
  contain a child element. FOV is a positive angle in 1/60000-degree units,
  strictly below 180 degrees; the profile keeps the native signed 32-bit
  bound as well.
- Projection stores the FOV override in the native shape model and keeps the
  existing camera/zoom/light-rig leaves independently available. The edit plan
  rechecks the full owner topology and token-splices only `camera/@fov`,
  preserving camera preset/zoom, light-rig state, root 3-D state, and all
  other package bytes.
- Missing, malformed, duplicate, extension-bearing, transformed,
  backdrop-bearing, unknown, default-only, or otherwise ambiguous scene state
  remains source-owned.

No source-free authoring, camera rotation editing, camera zoom editing,
light-rig editing, backdrop editing, relationship change, or full 3-D scene
reconstruction is included.
