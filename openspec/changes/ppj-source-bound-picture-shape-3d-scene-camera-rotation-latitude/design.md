## Context

The existing camera-rotation leaves use a strict scene reader that requires a
recognized camera, one child-free rotation node with canonical lat/lon/rev,
and a recognized light-rig pair. Picture scene preset, zoom, and FOV now use
the same source-bound approach.

## Goals / Non-Goals

**Goals:**

- Expose one canonical non-negative latitude in 60000ths-of-a-degree.
- Replace only the direct rotation `lat` token in the owning SlidePart.
- Preserve longitude, revolution, camera state, light rig, picture topology,
  and unrelated package parts.

**Non-Goals:**

- Partial rotation editing, rotation-node creation, or source-free authoring.
- Editing longitude/revolution, light-rig rotation, backdrop, or other scene
  state in this change.
- Reconstructing or normalizing unsupported markup.

## Decisions

- Reuse the existing full-camera-rotation reader and requested-value validator
  so picture and shape owners share the same canonical range and topology.
- Store the projected value in an optional `PresentationImage` field and reject
  it during new-picture authoring because it is source-bound state.
- Add picture dispatch to native proof/read and splice only `camera/rot/@lat`;
  broad scene serialization would risk changing source-owned state.

## Risks / Trade-offs

- [Risk] A rotation node with only latitude could be misinterpreted as a
  complete camera rotation. → Require canonical lat/lon/rev together.
- [Risk] Complex or extension-bearing scenes could be damaged by rewriting.
  → Keep the shared strict proof and fail closed.
