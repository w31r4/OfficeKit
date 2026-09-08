# Design

- Owner: `p:sp/p:spPr/a:scene3d/a:camera/a:rot/@lat`.
- Native leaf: `shape3dSceneCameraRotationLatitude60000`.
- The owner must be the single direct `a:scene3d` child of `a:spPr`. The
  scene must have exactly two direct children in schema order: an `a:camera`
  with one recognized `prst` attribute and exactly one bare `a:rot` child,
  followed by an `a:lightRig` with recognized `rig` and `dir` attributes and
  no children. The rotation must have exactly `lat`, `lon`, and `rev`
  attributes, all canonical non-negative 1/60000-degree tokens no greater than
  360 degrees. Camera zoom/FOV, light rotation, extensions, and backdrop state
  are not modeled.
- Projection stores only the camera rotation latitude in the native shape
  model. The edit plan rechecks the complete camera-rotation owner and
  token-splices only `rot/@lat`, preserving longitude, revolution, light-rig
  identity, root 3-D state, and all other package bytes.
- Missing, malformed, duplicate, extension-bearing, transformed,
  backdrop-bearing, unknown, default-only, or otherwise ambiguous scene state
  remains source-owned.

No source-free authoring, camera preset/zoom/FOV editing, light-rig editing,
longitude/revolution editing, backdrop editing, relationship change, or full
3-D scene reconstruction is included.
