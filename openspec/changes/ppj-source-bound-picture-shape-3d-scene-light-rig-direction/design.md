# Context

The existing scene-pair reader requires one direct scene with a recognized
camera, a child-free light rig with canonical `rig` and `dir`, and no unknown
attributes or children. Picture camera and light-rig preset leaves already
reuse this strict source-bound owner.

# Goals / Non-Goals

**Goals:**

- Expose one recognized DrawingML light-rig direction token.
- Replace only `lightRig/@dir` in the owning SlidePart.
- Preserve light-rig preset, camera state, picture topology, and unrelated
  package parts.

**Non-Goals:**

- Light-rig rotation, scene reconstruction, or source-free authoring.
- Editing unsupported effect/extension/relationship topology.

# Decisions

- Reuse `TryReadSceneLightRigDirection` and its requested-value validator so
  picture and shape owners share the same strict pair proof.
- Store the projected value in an additive `PresentationImage` field and
  reject it during new-picture authoring because it is source-bound state.
- Token-splice only the direct `lightRig/@dir` attribute.

# Risks / Trade-offs

- [Risk] A direction token without a valid camera/preset pair could be
  mistaken for a complete scene. → Require the existing strict scene pair.
- [Risk] Complex scene markup could be changed accidentally. → Fail closed
  unless the shared reader proves the exact owner.
