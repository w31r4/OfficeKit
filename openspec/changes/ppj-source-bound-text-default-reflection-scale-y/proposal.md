# Proposal: Add the paragraph default reflection vertical scale leaf

The remaining F-03 text-effect slice has a strict paragraph `defaultText`
reflection owner, but its vertical scale is still source-owned. Add one
source-bound PPJ native leaf so an imported `a:reflection/@sy` token can be
edited without rebuilding the effect list or touching another package part.

The field is deliberately narrow: it applies only to a direct full-span
paragraph default reflection whose only reflection transforms are the already
supported `sx`, `sy`, and `fadeDir`. Other transforms and malformed effect
graphs remain opaque or fail closed.
