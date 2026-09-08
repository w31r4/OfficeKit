# Design

## Owner and boundary

- Owner: `p:sp/p:spPr/a:custGeom/a:pathLst/a:path/a:cubicBezTo/a:pt[0]/@y`.
- PPJ surface: `shape.nativeRef.leaves[]`.
- Native leaf: `customGeometryPathCubicControl1Y`.
- `nativeLeafIndex` is the ordered path index and `textLeafIndex` is the
  ordered direct command index within that path.
- The geometry must pass the existing bounded custom-geometry recognition
  profile. The selected `cubicBezTo` and all three direct points must be
  attribute-only, and the first control-point y must be a canonical signed
  DrawingML coordinate; formula/reference-backed y values are not issued.

## Edit proof

The ordered path and command indexes plus the existing source/element/leaf
hashes identify the owner. Before editing, the codec rechecks the full path
topology, the selected direct `cubicBezTo` with exactly two control points and
one end point, the canonical numeric first-control-point y token, and the
requested changed bounded coordinate. The writer then replaces only the
selected `y` value token on the first point.

## Non-goals

- No public path-command array or command insertion/deletion.
- No line-to, move-to, quadratic, cubic second-control-point or end-point,
  first-control-point x, formula/reference, path property, or multi-path
  topology edits.
- No descendant rescale, 3-D effect, or unsupported/extension-bearing graph
  editing.
- No protocol field or Office wire version change.

## Evidence

The focused regression authors a custom geometry with move-to, line-to, and a
cubic command, removes embedded PPJ, projects the cubic first-control-point y
leaf, changes it, checks that only the owning SlidePart changed and the
package remains valid, then projects again to recover the new y coordinate and
unchanged control-point x and remaining point coordinates.
