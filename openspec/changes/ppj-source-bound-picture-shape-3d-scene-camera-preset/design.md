# Design

- Owner: `p:pic/p:spPr/a:scene3d/a:camera/@prst`.
- Native leaf: existing `shape3dSceneCameraPreset`.
- Source model: use the additive `PresentationImage.shape_3d_scene_camera_preset`
  field. It is source-projection-only and is not a source-free image
  authoring control.
- The picture owner must have exactly one direct `a:scene3d` with exactly two
  children in order: `a:camera` and `a:lightRig`. The camera must carry one
  recognized direct `prst` attribute and may contain only the existing helper's
  canonical optional camera `zoom`/`fov` attributes; the light rig must remain
  a recognized `rig`/`dir` pair owner. Unknown, malformed, duplicate,
  extension-bearing, extra-child, transformed, or other complex scene state
  remains source-owned.
- Picture projection stores the camera preset token, issues the existing
  native leaf, and keeps the image relationship, crop, mask, effects, and
  remaining 3-D XML untouched. The edit plan re-proves the picture scene and
  token-splices only `a:camera/@prst`.

No source-free authoring, geometry normalization, relationship changes,
effect editing, scene reconstruction, or descendant transform is included.
