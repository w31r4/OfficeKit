# Design

## Owner and boundary

- Owner: `p:sp/p:spPr/a:custGeom/a:pathLst/a:path/a:arcTo/@stAng`.
- PPJ surface: `shape.nativeRef.leaves[]`.
- Native leaf: `customGeometryPathArcStartAngle60000`.
- `nativeLeafIndex` is the ordered path index and `textLeafIndex` is the
  ordered direct command index within that path.
- The geometry must pass the existing bounded custom-geometry recognition
  profile. The selected arc must have only the four standard arc attributes,
  no child content, and a canonical literal start angle in the signed
  one-turn range; reference-backed sibling arc values remain source-owned.

## Edit proof

The ordered path and command indexes plus the existing source/element/leaf
hashes identify the owner. Before editing, the codec rechecks the recognized
path, selected direct `arcTo`, standard attribute-only topology, canonical
one-turn numeric `stAng`, and the requested changed bounded angle. The writer
then replaces only the selected `stAng` value token.

## Non-goals

- No public path-command array or command insertion/deletion.
- No arc radii, sweep angle, point coordinates, formula/reference, path
  property, or multi-path topology edits.
- No angle normalization, descendant rescale, 3-D effect, or unsupported/
  extension-bearing graph editing.
- No protocol field or Office wire version change.

## Evidence

The focused regression authors a custom geometry containing a move-to, line-to,
and arc command, removes embedded PPJ, projects the arc start-angle leaf,
changes it, checks that only the owning SlidePart changed and the package
remains valid, then projects again to recover the new angle and unchanged arc
attributes.
