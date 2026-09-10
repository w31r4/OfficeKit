# Proposal: Add the paragraph default reflection horizontal scale leaf

The remaining F-03 text-effect slice has a strict paragraph `defaultText`
reflection owner, but its horizontal scale is still source-owned. Add one
source-bound PPJ native leaf so an imported `a:reflection/@sx` token can be
edited without rebuilding the effect list or touching another package part.

The field is deliberately narrow: it applies only to a direct full-span
paragraph default reflection whose only reflection transform is `sx` (plus the
already supported `fadeDir`). Other transforms and malformed effect graphs
remain opaque or fail closed.
