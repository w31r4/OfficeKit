# Design

## PPJ contract

- Native leaf: `textDefaultReflectionSkewX`.
- PPJ location: `paragraph.style.defaultText.reflection.skewX`.
- Unit: signed degrees represented by the native `kx` integer at 1/60000
  degree precision.
- Accepted raw token: canonical signed Int32 whose rounded value is strictly
  between -5,400,000 and 5,400,000.

## Source-bound boundary

Projection emits the leaf only when the paragraph `a:defRPr` effect list has a
single strict full-span reflection and contains `kx`. The owner may also
contain the existing `sx`, `sy`, and `fadeDir`; `ky`, `algn`, `rotWithShape`,
unknown attributes or children, duplicate lists/effects, and DAGs remain
source-owned.

An edit proves the paragraph index and expected raw `kx` token, then replaces
only `a:reflection/@kx` in the owning SlidePart. All other reflection
attributes, sibling effects, text topology, and ZIP members remain unchanged.

## Evidence

Extend the source-bound text-effect fixture with a signed skew leaf. The
experiment checks projection, a `-12.5 -> 20` degree edit, native XML validity,
changed-part scope, second projection, and a `ky` rejection. No PowerPoint
host-rendering claim is made.
