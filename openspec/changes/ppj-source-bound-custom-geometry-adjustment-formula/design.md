# Design

## Owner and boundary

- Owner: `p:sp/p:spPr/a:custGeom/a:avLst/a:gd/@fmla`.
- PPJ surface: `shape.nativeRef.leaves[]`.
- Native leaf: `customGeometryAdjustmentFormula`.
- The direct `avLst` must be unique, attribute-only, contain named guides, and
  belong to an otherwise recognized bounded custom-geometry graph.
- The selected formula must be canonical under the existing guide parser and
  must not be the literal `val N` form already represented by
  `customGeometryAdjustment`.

## Edit proof

The ordered adjustment index plus the existing source/element/leaf hashes
identify the owner. Before the edit, the codec rechecks the direct list,
attributes, formula canonicalization, shape extents, and full graph profile.
The requested formula is installed into a cloned geometry and accepted only if
the existing formula codec parses and evaluates the complete graph and returns
the exact requested canonical string. The writer then replaces only the
selected `fmla` value token.

## Non-goals

- No public guide array or formula-language expansion.
- No edits to adjustment names/order, calculated guides, handles, connection
  sites, text rectangles, paths, extensions, or descendant/group topology.
- No partial or opaque image-fill geometry formula editing.
- No protocol field or Office wire version change.

## Evidence

The focused regression authors a custom geometry containing a calculated
adjustment, removes embedded PPJ, projects the formula leaf, changes only the
adjustment formula, verifies a SlidePart-only change and Open XML validity, then
projects again to recover the new formula. Other geometry XML and ZIP parts
remain unchanged.
