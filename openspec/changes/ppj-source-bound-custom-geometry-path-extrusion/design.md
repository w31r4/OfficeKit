## Context

See `proposal.md` for the motivation. The custom-geometry codec already
recognizes an optional direct path extrusion boolean and preserves its absence.
Generic native leaves are projected only when source identity and a narrow XML
token-splice can be proved.

## Goals / Non-Goals

**Goals:**

- Expose each recognized custom path's explicit `extrusionOk` token as a
  path-indexed boolean leaf.
- Preserve viewport attributes, fill/stroke properties, commands, sibling
  paths, and source ownership boundaries while editing only `@extrusionOk`.
- Demonstrate authored export, embedded-PPJ removal, import projection,
  source-bound edit, and second projection.

**Non-Goals:**

- Editing width, height, fill, stroke, commands, path order, or path-list
  topology.
- Materializing omitted extrusion permission or changing 3D shape effects.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use path order as native identity.** `NativeLeafIndex` follows the direct
   `a:path` order in `a:pathLst`, matching the typed path array and avoiding a
   synthetic path identifier.

2. **Expose the explicit boolean only.** A direct canonical boolean token
   becomes a PPJ boolean. Requested PPJ values are normalized to the same
   boolean and rejected when unchanged or non-canonical.

3. **Preserve omission.** A missing `@extrusionOk` keeps the native default and
   does not receive a leaf, so an edit cannot silently materialize an
   attribute.

4. **Prove the minimal XML owner.** The proof checks one direct custom geometry
   path list, the path index, supported path structure, and an explicit
   extrusion boolean. Compilation targets that path's `extrusionOk` attribute
   and uses the existing source element hash plus raw-value precondition.

## Risks / Trade-offs

- [Risk] Toggling extrusion permission may expose or hide a 3D extrusion effect
  whose full graph remains opaque. → Mitigation: edit only the direct boolean
  token and preserve all other path and shape properties.
- [Risk] A future producer may write an unexpected lexical boolean. →
  Mitigation: canonical source values are proved before projection/editing;
  malformed or unsupported profiles remain source-owned.

## Migration Plan

No migration is required. Existing PPJ programs remain valid; the new leaf is
only emitted for newly projected sources that satisfy the bounded profile.

## Open Questions

None.
