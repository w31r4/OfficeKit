# Design: source-bound picture shadow rotation

## Owner

The owner is one picture shape-properties graph:

```text
p:pic/a:spPr/a:effectLst/a:outerShdw/@rotWithShape
```

The existing strict shadow reader preserves the optional `rotWithShape`
attribute. This slice exposes it only when the attribute is present and uses
the canonical `0`/`1` boolean form; the picture payload, mask, border, shadow
color, geometry and other effect topology remain source-owned.

## Safety boundary

Missing `rotWithShape`, multiple effect lists, multiple effects, non-direct or
extension children, malformed topology, and relationship changes remain
source-owned or fail closed. No source-bound insertion, removal, or reordering
is allowed.

## Edit proof

The leaf kind is `imageShadowRotateWithShape` and its PPJ location is
`image.shadow.rotateWithShape`. The edit plan re-proves the selected picture,
direct picture properties, bounded effect list, direct `a:outerShdw`, and
explicit `rotWithShape` attribute before token-splicing only that attribute in
the owning `SlidePart`. No protobuf or wire-version change is needed. The
focused source-bound picture fixture adds the explicit authored boolean,
changes the leaf, verifies SlidePart-only mutation and package validity, and
reprojects the value.
