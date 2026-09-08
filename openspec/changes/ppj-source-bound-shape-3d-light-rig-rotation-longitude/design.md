# Design

- Owner: `p:sp/p:spPr/a:scene3d/a:lightRig/a:rot/@lon`.
- Native leaf: `shape3dSceneLightRigRotationLongitude60000`.
- The owner must be the single direct `a:scene3d` child of `a:spPr`. The
  scene must have exactly two direct children in schema order: a bare
  `a:camera` with one recognized `prst`, followed by an `a:lightRig` with
  recognized `rig` and `dir` attributes and exactly one bare `a:rot` child.
  The rotation must have exactly `lat`, `lon`, and `rev` attributes, all
  canonical non-negative 1/60000-degree tokens no greater than 360 degrees.
  Camera overrides, nested children, extensions, and backdrop state are not
  modeled.
- Projection stores only the rotation longitude in the native shape model. The
  edit plan rechecks the complete rotation owner and token-splices only
  `rot/@lon`, preserving latitude, revolution, camera/light-rig identity, root
  3-D state, and all other package bytes.
- Missing, malformed, duplicate, extension-bearing, transformed,
  backdrop-bearing, unknown, default-only, or otherwise ambiguous scene state
  remains source-owned.

No source-free authoring, camera editing, light-rig preset/direction editing,
latitude/revolution editing, backdrop editing, relationship change, or full
3-D scene reconstruction is included.
