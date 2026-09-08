# Design

- Owner: `p:sp/p:spPr/a:scene3d/a:lightRig/@dir`.
- Native leaf: `shape3dSceneLightRigDirection`.
- The owner must be the single direct `a:scene3d` child of `a:spPr`. The
  `a:scene3d` owner must have exactly two direct children in schema order: a
  bare `a:camera` with only a recognized `prst`, and a bare `a:lightRig` with
  recognized `rig` and `dir` attributes. Neither child may contain a child
  element. Camera rotation/FOV/zoom, light rotation, backdrop, and extension
  state are not modeled.
- Projection stores only the bounded light-rig direction in the native shape
  model and emits one leaf. The edit plan rechecks the full owner topology and
  token-splices only `lightRig/@dir`, preserving the camera, light-rig preset,
  root 3-D state, and all other package bytes.
- Missing, malformed, duplicate, extension-bearing, transformed,
  backdrop-bearing, unknown, or otherwise ambiguous scene state remains
  source-owned.

No source-free authoring, camera editing, light-rig preset editing, light
rotation editing, backdrop editing, relationship change, or full 3-D scene
reconstruction is included.
