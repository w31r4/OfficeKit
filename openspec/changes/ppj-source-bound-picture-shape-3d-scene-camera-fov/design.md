## Context

The shape 3-D scene FOV leaf already has a strict scene-pair reader and a
changed-value validator. Picture scene preset and zoom now use the same
source-bound pattern. This change applies it to the direct picture camera FOV
attribute.

## Goals / Non-Goals

**Goals:**

- Expose one explicit canonical positive camera FOV below 180 degrees in
  60000ths-of-a-degree.
- Replace only a changed `a:camera/@fov` token in the owning SlidePart.
- Preserve the picture asset topology and all unrelated scene/package markup.

**Non-Goals:**

- Source-free picture scene authoring or default FOV synthesis.
- Editing camera preset/zoom/rotation, light-rig state, backdrop, or arbitrary
  scene topology.
- Reconstructing or normalizing unsupported markup.

## Decisions

- Reuse the existing scene-pair FOV reader and requested-value validator so
  picture and shape owners share the same recognized camera/light-rig and FOV
  bounds.
- Add an optional `PresentationImage` field as a source-projection-only value;
  new pictures reject it rather than silently dropping source-bound state.
- Extend native proof/read dispatch to `p:pic` and perform one direct XML
  attribute splice. A broad scene rewrite would unnecessarily take ownership
  of unrelated state.

## Risks / Trade-offs

- [Risk] Missing FOV may represent an Office default rather than an editable
  source token. → Expose only explicit canonical FOV and keep omission opaque.
- [Risk] Complex scene markup may be corrupted by reconstruction. → Require the
  shared strict scene proof and fail closed.
