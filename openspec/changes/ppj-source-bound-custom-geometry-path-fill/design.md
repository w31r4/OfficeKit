## Context

See `proposal.md` for the motivation. The custom-geometry codec already
recognizes direct path fill values `norm` and `none`, while preserving omitted
fill as the default. Generic native leaves are projected only when source
identity and a narrow XML token-splice can be proved.

## Goals / Non-Goals

**Goals:**

- Expose each recognized custom path's explicit `norm`/`none` fill token as a
  path-indexed boolean leaf.
- Preserve viewport attributes, stroke/extrusion properties, commands, sibling
  paths, and source ownership boundaries while editing only `@fill`.
- Demonstrate authored export, embedded-PPJ removal, import projection,
  source-bound edit, and second projection.

**Non-Goals:**

- Editing width, height, stroke, extrusion, commands, path order, or path-list
  topology.
- Materializing omitted fill, accepting other DrawingML fill modes, or changing
  formula-driven geometry.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use path order as native identity.** `NativeLeafIndex` follows the direct
   `a:path` order in `a:pathLst`, matching the typed path array and avoiding a
   synthetic path identifier.

2. **Map only the recognized modes.** Direct `norm` becomes `true` and direct
   `none` becomes `false`. The requested PPJ value is normalized to the same
   boolean and is rejected when unchanged or non-canonical.

3. **Preserve omission.** A missing `@fill` keeps the native default and does
   not receive a leaf, so an edit cannot silently materialize an attribute.

4. **Prove the minimal XML owner.** The proof checks one direct custom geometry
   path list, the path index, supported path structure, and explicit fill mode.
   Compilation targets that path's `fill` attribute and uses the existing
   source element hash plus raw-value precondition.

## Risks / Trade-offs

- [Risk] Switching fill mode changes whether the path contributes a filled
  region. → Mitigation: expose only the two codec-recognized direct modes and
  keep all path commands and other paint properties untouched.
- [Risk] A future DrawingML extension may add a different path fill value.
  → Mitigation: the proof rejects every value outside `norm`/`none` and keeps
  omitted/unknown profiles source-owned.

## Migration Plan

No migration is required. Existing PPJ programs remain valid; the new leaf is
only emitted for newly projected sources that satisfy the bounded profile.

## Open Questions

None.
