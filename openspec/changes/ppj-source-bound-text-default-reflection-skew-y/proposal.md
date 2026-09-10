# Proposal: Add the paragraph default reflection vertical skew leaf

The F-03 text-effect profile now exposes paragraph `defaultText` reflection
scale and horizontal-skew fields, but vertical skew remains opaque. Add one
narrow source-bound PPJ native leaf so an imported `a:reflection/@ky` token can
be edited without rebuilding the effect list or touching another package part.

The field applies only to a direct full-span reflection whose allowed
transforms are the already modelled `sx`, `sy`, `fadeDir`, `kx`, and the new
`ky`. Alignment, rotation, unknown content, and malformed graphs remain
source-owned or fail closed.
