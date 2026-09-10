# Design

## PPJ contract

- Native leaf: `textDefaultReflectionScaleY`.
- PPJ location: `paragraph.style.defaultText.reflection.scaleY`.
- Unit: signed ratio, represented by the native `sy` integer at 1/100000
  precision.
- Accepted raw token: canonical signed Int32 (`-2147483648..2147483647`).

## Source-bound boundary

Projection emits the leaf only when the paragraph `a:defRPr` effect list has a
single strict full-span reflection owner and contains `sy`. The owner may also
contain the already modelled `sx` and `fadeDir`; `kx`, `ky`, `algn`,
`rotWithShape`, unknown attributes, children, duplicate lists/effects and DAGs
remain source-owned.

An edit proves the paragraph index and expected raw `sy` token, then replaces
only `a:reflection/@sy` in the owning SlidePart. The package compiler keeps all
other XML and ZIP members unchanged.

## Evidence

Extend the existing source-bound text-effect fixture with one scale leaf. The
experiment checks the authored/imported value, native leaf, changed-part scope,
Open XML validity, re-projection, and an unsupported-transform rejection. No
PowerPoint host-rendering claim is made.
