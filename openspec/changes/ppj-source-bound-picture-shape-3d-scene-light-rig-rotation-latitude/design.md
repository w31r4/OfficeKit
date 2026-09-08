# Context

The existing light-rig rotation reader requires one direct scene with a bare
recognized camera, a canonical light-rig pair, and one child-free rotation
with literal non-negative latitude/longitude/revolution values. Picture scene
leaves already reuse the strict two-child owner.

# Goals / Non-Goals

**Goals:**

- Expose one canonical signed `a:rot/@lat` angle in 60000ths of a degree.
- Replace only `lightRig/a:rot/@lat` in the owning SlidePart.
- Preserve longitude, revolution, camera state, light-rig state, picture
  topology, and unrelated package parts.

**Non-Goals:**

- Partial rotation, scene reconstruction, or source-free authoring.
- Editing unsupported effect/extension/relationship topology.

# Decisions

- Reuse `TryReadSceneLightRigRotationLatitude` and its requested-value
  validator so picture and shape owners share the same complete-rotation proof.
- Store the projected value in an optional `PresentationImage` field and
  reject it during new-picture authoring because it is source-bound state.
- Token-splice only the direct `lightRig/a:rot/@lat` attribute.

# Risks / Trade-offs

- [Risk] A latitude token without a complete rotation owner could be mistaken
  for a safe scene. → Require the shared full-rotation reader.
- [Risk] A partial or extended rotation could be normalized accidentally. →
  Fail closed unless the exact owner is proven.
