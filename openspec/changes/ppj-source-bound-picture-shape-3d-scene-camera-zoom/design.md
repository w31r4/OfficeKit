## Context

The existing shape 3-D scene camera zoom leaf uses a shared strict scene-pair
validator and token splice. Picture 3-D scalar and color leaves already use an
additive `PresentationImage` source-projection field and the same native leaf
surface. This change follows those boundaries for the direct `p:pic` owner.

## Goals / Non-Goals

**Goals:**

- Expose one canonical non-negative `a:camera/@zoom` value from a strict
  picture scene.
- Accept a changed valid zoom token and replace only that attribute in the
  owning SlidePart.
- Keep the picture asset relationship, crop/mask/effects, camera preset,
  light rig, other scene state, and unrelated package parts unchanged.

**Non-Goals:**

- Authoring a new source-free picture scene or synthesizing omitted/default
  zoom.
- Editing camera preset/FOV/rotation, light-rig state, backdrop, or arbitrary
  scene topology in this change.
- Normalizing, reconstructing, or flattening unsupported picture markup.

## Decisions

- Reuse the existing scene-pair reader and requested-value validator. This
  keeps shape and picture owners on one definition of recognized camera,
  canonical non-negative zoom, and recognized light-rig state; a separate
  picture parser would risk drifting boundaries.
- Add the field to `PresentationImage` rather than changing the wire version.
  The value exists only when projection proves the source owner and is rejected
  for new source-free picture authoring.
- Extend the existing native leaf to `image.nativeRef.leaves[]` and make the
  edit-plan owner dispatch accept `p:pic`. The edit remains a single attribute
  splice, so the surrounding XML stays source-owned.

## Risks / Trade-offs

- [Risk] A camera with omitted zoom cannot distinguish an authored default from
  an absent value. → Keep omitted/default zoom opaque and expose only an
  explicit canonical token.
- [Risk] An incomplete or extension-bearing scene could be damaged by a broad
  rewrite. → Require the shared strict scene proof and fail closed otherwise.
