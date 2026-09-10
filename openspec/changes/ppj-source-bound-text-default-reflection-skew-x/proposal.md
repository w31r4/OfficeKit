# Proposal: Add the paragraph default reflection horizontal skew leaf

The F-03 text-effect profile already owns paragraph `defaultText` reflection
scale and fade fields, but horizontal skew remains opaque. Add one narrow
source-bound PPJ native leaf so an imported `a:reflection/@kx` token can be
edited without rebuilding the effect list or touching another package part.

The field applies only to a direct full-span reflection whose allowed
transforms are the already modelled `sx`, `sy`, `fadeDir`, and the new `kx`.
`ky`, alignment, rotation, unknown content, and malformed graphs remain
source-owned or fail closed.
