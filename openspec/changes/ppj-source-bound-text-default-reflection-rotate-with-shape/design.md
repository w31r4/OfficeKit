# Design

## Owner and representation

Use the existing direct paragraph default-run owner:
`p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:reflection`. The PPJ field is
`paragraph.style.defaultText.reflection.rotateWithShape`; its native leaf kind is
`textDefaultReflectionRotateWithShape`, a boolean represented by the existing
canonical `rotWithShape` token (`0` or `1`). The reflection must remain a strict
full-span owner, with the already-modelled transform and alignment attributes
only; no relationship, protobuf, or wire version changes are needed.

## Projection

After the normal strict owner proof succeeds, issue one paragraph-index-bound
leaf when `rotWithShape` is explicitly present. Preserve the parsed reflection
object and all other source attributes. Missing or malformed flags, unknown
attributes/children, duplicate lists/effects, DAGs, or non-full-span positions
keep the source-owned boundary and do not issue the leaf.

## Source-bound edit

Extend the existing paragraph default reflection edit-plan whitelist, expected
value proof, and XML attribute map. Validate both expected and requested values
as changed canonical booleans, then replace only the `rotWithShape` value in the
existing reflection start tag. Keep changed-part and source hash proofs unchanged.

## Regression

Extend the existing focused source-bound text-effect fixture: set an explicit
source `rotWithShape=1`, assert the PPJ boolean and native leaf, edit it to `0`,
assert only the owning SlidePart changes and Open XML remains valid, and assert a
second projection returns `false`. Move the old unsupported-transform negative
case to a non-full-span position because `rotWithShape` is now modelled; the
existing unknown-attribute and child cases continue to prove fail-closed
behavior.
