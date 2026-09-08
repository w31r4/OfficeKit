# Context

The existing backdrop reader requires one direct scene with a bare recognized
camera, a canonical light-rig pair, one backdrop, and complete bare anchor,
normal, and up vectors with signed coordinate literals. Picture scene leaves
should reuse that strict owner.

# Goals / Non-Goals

**Goals:**

- Expose one canonical signed `a:backdrop/a:up/@dz` coordinate in EMU.
- Replace only `backdrop/a:up/@dz` in the owning SlidePart.
- Preserve the anchor, the other vector coordinates, camera/light-rig state,
  picture topology, and unrelated package parts.

**Non-Goals:**

- Partial backdrop, scene reconstruction, or source-free authoring.
- Editing unsupported effect/extension/relationship topology.

# Decisions

- Reuse `TryReadSceneBackdropUpDz` and its requested-value validator so picture
  and shape owners share the same complete-backdrop proof.
- Store the projected value in an optional `PresentationImage` field and
  reject it during new-picture authoring because it is source-bound state.
- Token-splice only the direct `backdrop/a:up/@dz` attribute.

# Risks / Trade-offs

- [Risk] An up-vector Z token without a complete backdrop owner could be
  mistaken for a safe scene. → Require the shared complete-backdrop reader.
- [Risk] A partial or extended backdrop could be normalized accidentally. →
  Fail closed unless the exact owner is proven.
