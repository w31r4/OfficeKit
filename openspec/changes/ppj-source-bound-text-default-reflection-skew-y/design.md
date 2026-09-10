# Design

## PPJ contract

- Native leaf: `textDefaultReflectionSkewY`.
- PPJ location: `paragraph.style.defaultText.reflection.skewY`.
- Unit: signed degrees represented by the native `ky` integer at 1/60000
  degree precision.
- Accepted raw token: canonical signed Int32 whose rounded value is strictly
  between -5,400,000 and 5,400,000.

## Source-bound boundary

The native projection may retain a parsed scalar for inspection, but the leaf
is editable only when the paragraph `a:defRPr` effect list has a single strict
full-span reflection and contains `ky`. The owner may also contain the existing
`sx`, `sy`, `fadeDir`, and `kx`; `algn`, `rotWithShape`, unknown attributes or
children, duplicate lists/effects, and DAGs remain source-owned and fail closed
for edits.

An edit proves the paragraph index and expected raw `ky` token, then replaces
only `a:reflection/@ky` in the owning SlidePart. All other reflection
attributes, sibling effects, text topology, and ZIP members remain unchanged.

## Evidence

Extend the source-bound text-effect fixture with a signed vertical-skew leaf.
The experiment checks projection, a `8.25 -> -20` degree edit, native XML
validity, changed-part scope, second projection, and an alignment rejection.
No PowerPoint host-rendering claim is made.
