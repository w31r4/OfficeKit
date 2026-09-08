## Context

See `proposal.md` for the motivation. The existing custom-geometry codec
recognizes `a:path` command topology and carries its optional height as a
DrawingML coordinate. Generic native leaves are projected only when the source
identity and a narrow XML token-splice can be proved.

## Goals / Non-Goals

**Goals:**

- Expose each recognized custom path's existing positive direct `@h` as a
  path-indexed native leaf.
- Preserve the path's other attributes, commands, sibling paths, and source
  ownership boundaries while editing only `@h`.
- Demonstrate authored export, embedded-PPJ removal, import projection,
  source-bound edit, and second projection.

**Non-Goals:**

- Editing `@w`, fill/stroke/extrusion flags, commands, path order, or path-list
  topology.
- Materializing an omitted or zero height, changing formula-driven coordinates,
  or accepting extension/unknown path children.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use path order as native identity.** `NativeLeafIndex` follows the direct
   `a:path` order in `a:pathLst`, matching the typed path array and avoiding a
   synthetic path identifier.

2. **Expose only positive canonical heights.** A direct numeric height must be
   a positive integer no larger than `int.MaxValue`; omission and zero retain
   the codec's existing semantics and are not made editable. The requested
   value must use the same canonical decimal form and remain positive and
   bounded.

3. **Keep the unit explicit.** The leaf is named without an `Emu` suffix and
   its registry unit is DrawingML path-coordinate units. `a:path/@h` defines
   the path viewport coordinate system, not the shape's EMU frame.

4. **Prove the minimal XML owner.** The proof checks a single direct custom
   geometry/path list, an in-range recognized path index, a positive height, and
   known path command children. Compilation targets that path's `h` attribute
   and lets the existing expected-value and element-hash checks reject stale or
   structurally changed sources.

## Risks / Trade-offs

- [Risk] Changing a path viewport height can alter the rendered geometry because
  command coordinates are interpreted in that viewport. → Mitigation: expose
  only an explicitly existing height, label the unit accurately, and preserve
  every command while keeping broader topology source-owned.
- [Risk] A future DrawingML extension may appear as a path child or attribute.
  → Mitigation: the proof fails closed unless the direct path structure is
  recognized and limited to the supported command/attribute set.

## Migration Plan

No migration is required. Existing PPJ programs remain valid; the new leaf is
only emitted for newly projected sources that satisfy the bounded profile.

## Open Questions

None.
