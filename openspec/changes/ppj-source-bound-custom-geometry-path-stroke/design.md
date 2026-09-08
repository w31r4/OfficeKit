## Context

See `proposal.md` for the motivation. The custom-geometry codec already
recognizes direct path stroke values as an optional boolean, while preserving
omitted stroke as the native default. Generic native leaves are projected only
when source identity and a narrow XML token-splice can be proved.

## Goals / Non-Goals

**Goals:**

- Expose each recognized custom path's explicit stroke token as a path-indexed
  boolean leaf.
- Preserve viewport attributes, fill/extrusion properties, commands, sibling
  paths, and source ownership boundaries while editing only `@stroke`.
- Demonstrate authored export, embedded-PPJ removal, import projection,
  source-bound edit, and second projection.

**Non-Goals:**

- Editing width, height, fill, extrusion, commands, path order, or path-list
  topology.
- Materializing omitted stroke or changing formula-driven geometry.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use path order as native identity.** `NativeLeafIndex` follows the direct
   `a:path` order in `a:pathLst`, matching the typed path array and avoiding a
   synthetic path identifier.

2. **Expose the explicit boolean only.** A direct `stroke="1"`/`"0"` token
   becomes `true`/`false`. Requested PPJ values are normalized to the same
   boolean and rejected when unchanged or non-canonical.

3. **Preserve omission.** A missing `@stroke` keeps the native default and does
   not receive a leaf, so an edit cannot silently materialize an attribute.

4. **Prove the minimal XML owner.** The proof checks one direct custom geometry
   path list, the path index, supported path structure, and an explicit boolean
   stroke. Compilation targets that path's `stroke` attribute and uses the
   existing source element hash plus raw-value precondition.

## Risks / Trade-offs

- [Risk] A stroke toggle changes the visible outline while leaving the rest of
  the path opaque. → Mitigation: edit only the direct boolean token and keep
  all commands and other paint properties untouched.
- [Risk] A future DrawingML producer may write an unexpected lexical boolean.
  → Mitigation: canonical source values are proved before projection/editing;
  malformed or unsupported profiles remain source-owned.

## Migration Plan

No migration is required. Existing PPJ programs remain valid; the new leaf is
only emitted for newly projected sources that satisfy the bounded profile.

## Open Questions

None.
