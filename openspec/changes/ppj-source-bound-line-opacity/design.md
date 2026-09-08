# Design

## Owner and boundary

- Owners: `p:sp/a:spPr/a:ln/a:solidFill/{a:srgbClr|a:schemeClr}/a:alpha/@val`
  and the equivalent `p:cxnSp/a:spPr/a:ln/...` path.
- PPJ semantic location: `shape.line.opacity` or `connector.line.opacity`.
- Native leaf: `lineOpacityThousandthPercent`.
- The exposed value is the existing canonical integer `0..100000` in
  DrawingML thousandths of a percent.
- The owner must have one direct outline, one direct solid RGB or theme paint,
  and one existing direct alpha token. Missing alpha is not materialized by
  this change; no-fill, gradient, image, compound, inherited, transformed,
  extension-bearing, or otherwise ambiguous outlines remain source-owned.
- Editing replaces only the existing `a:alpha/@val` token. It preserves line
  color, width, dash, cap, join, endpoints, shape geometry, and all unrelated
  package parts.

## Evidence

The focused regression covers one ordinary shape and one connector. It removes
the embedded PPJ, projects the native leaf, edits the alpha, verifies that only
the owning SlidePart changes, validates Open XML, and confirms the new opacity
after a second projection. A missing-alpha negative fixture confirms that the
capability is not invented.
