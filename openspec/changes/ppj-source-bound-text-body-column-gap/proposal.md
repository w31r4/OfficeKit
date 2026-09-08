# Proposal: expose source-bound text-body column gap as a PPJ leaf

Expose the existing direct text-container `columnGap` value as an independent
native leaf for imported text shapes. The leaf lets a caller change only the
existing DrawingML `a:bodyPr/@spcCol` token while preserving the text body,
paragraph/run topology, and the rest of the SlidePart.

This is a scalar continuation of the existing bounded `textBoxStyle.columnGap`
profile. Missing body properties, inherited values, malformed values, unknown
attributes/children, and table-cell style operations remain on their existing
typed source-bound paths or fail closed.
