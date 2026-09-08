# Design

## Owner and boundary

- Owner: `p:sp/p:spPr/a:custGeom/a:pathLst/a:path/a:moveTo/a:pt/@x`.
- PPJ surface: `shape.nativeRef.leaves[]`.
- Native leaf: `customGeometryPathMoveToX`.
- `nativeLeafIndex` is the ordered path index and `textLeafIndex` is the
  ordered direct command index within that path.
- The geometry must pass the existing bounded custom-geometry recognition
  profile. The selected `moveTo` and `pt` must be attribute-only, and `x` must
  be a canonical signed DrawingML coordinate; formula/reference-backed x
  values are not issued.

## Edit proof

The ordered path and command indexes plus the existing source/element/leaf
hashes identify the owner. Before editing, the codec rechecks the full path
topology, the selected direct `moveTo/pt`, the canonical numeric x token, and
the requested changed bounded coordinate. The writer then replaces only the
selected `x` value token.

## Non-goals

- No public path-command array or command insertion/deletion.
- No line-to, quadratic/cubic control-point, arc, formula/reference, path
  property, or multi-path topology edits.
- No descendant rescale, 3-D effect, or unsupported/extension-bearing graph
  editing.
- No protocol field or Office wire version change.

## Evidence

The focused regression authors a custom geometry with a move-to and line-to,
removes embedded PPJ, projects the move-to x leaf, changes it, checks that only
the owning SlidePart changed and the package remains valid, then projects again
to recover the new x coordinate and unchanged paired y.
