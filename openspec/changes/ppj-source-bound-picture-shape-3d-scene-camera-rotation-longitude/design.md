## Context

The existing camera-rotation leaves use a strict scene reader that requires a
recognized camera, one child-free rotation node with canonical lat/lon/rev,
and a recognized light-rig pair. Picture scene preset, zoom, FOV, and latitude
now use the same source-bound approach.

## Goals / Non-Goals

**Goals:**

- Expose one canonical non-negative longitude in 60000ths-of-a-degree.
- Replace only the direct rotation `lon` token in the owning SlidePart.
- Preserve latitude, revolution, camera state, light rig, picture topology,
  and unrelated package parts.

**Non-Goals:**

- Partial rotation editing, rotation-node creation, or source-free authoring.
- Editing latitude/revolution, light-rig rotation, backdrop, or other scene
  state in this change.
- Reconstructing or normalizing unsupported markup.

## Decisions

- Reuse the existing full-camera-rotation reader and requested-value validator
  so picture and shape owners share the same canonical range and topology.
- Store the projected value in an optional `PresentationImage` field and reject
  it during new-picture authoring because it is source-bound state.
- Add picture dispatch to native proof/read and splice only `camera/rot/@lon`;
  broad scene serialization would risk changing source-owned state.

## Risks / Trade-offs

- [Risk] A rotation node with only longitude could be misinterpreted as a
  complete camera rotation. → Require canonical lat/lon/rev together.
- [Risk] Complex or extension-bearing scenes could be damaged by rewriting.
  → Keep the shared strict proof and fail closed.
