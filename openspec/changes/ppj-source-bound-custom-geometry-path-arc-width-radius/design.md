# Design

## Owner and boundary

- Owner: `p:sp/p:spPr/a:custGeom/a:pathLst/a:path/a:arcTo/@wR`.
- PPJ surface: `shape.nativeRef.leaves[]`.
- Native leaf: `customGeometryPathArcWidthRadius`.
- `nativeLeafIndex` is the ordered path index and `textLeafIndex` is the
  ordered direct command index within that path.
- The geometry must pass the existing bounded custom-geometry recognition
  profile. The selected arc must have only the four standard arc attributes,
  no child content, and a positive canonical literal `wR`; reference-backed
  sibling arc values remain source-owned.

## Edit proof

The ordered path and command indexes plus the existing source/element/leaf
hashes identify the owner. Before editing, the codec rechecks the recognized
path, selected direct `arcTo`, standard attribute-only topology, canonical
positive numeric `wR`, and the requested changed bounded radius. The writer
then replaces only the selected `wR` value token.

## Non-goals

- No public path-command array or command insertion/deletion.
- No arc height radius, start angle, sweep angle, point coordinates, formula/
  reference, path property, or multi-path topology edits.
- No descendant rescale, 3-D effect, or unsupported/extension-bearing graph
  editing.
- No protocol field or Office wire version change.

## Evidence

The focused regression authors a custom geometry containing a move-to, line-to,
and arc command, removes embedded PPJ, projects the arc width-radius leaf,
changes it, checks that only the owning SlidePart changed and the package
remains valid, then projects again to recover the new radius and unchanged arc
attributes.
