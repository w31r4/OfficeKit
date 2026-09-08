# Design

- Owner: `p:sp/p:spPr/a:scene3d/a:camera/@prst`.
- Native leaf: `shape3dSceneCameraPreset`.
- The owner must be the single direct `a:scene3d` child of `a:spPr`. The
  `a:scene3d` owner must have exactly two
  direct children in schema order: a bare `a:camera` with only `prst`, and a
  bare `a:lightRig` with only bounded `rig` and `dir` attributes. No camera
  rotation, FOV, zoom, light rotation, backdrop, or extension is modeled.
  Direct `a:sp3d` root attributes remain limited to the existing scalar
  attributes and are preserved.
- Projection stores only the bounded camera preset token in the native shape
  model and emits one leaf. The edit plan rechecks the full owner topology and
  token-splices only `camera/@prst`, preserving the scene, light rig, root
  3-D state, and all other package bytes.
- Missing, malformed, duplicate, extension-bearing, transformed,
  backdrop-bearing, unknown, or otherwise ambiguous scene state remains
  source-owned.

No source-free authoring, camera transform editing, light-rig editing,
backdrop editing, relationship change, or full 3-D scene reconstruction is
included.
