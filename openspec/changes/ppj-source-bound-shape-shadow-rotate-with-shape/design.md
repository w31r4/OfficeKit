# Design: source-bound ordinary-shape shadow rotation

## Owner

The owner is one ordinary shape or line shape-properties graph:

```text
p:sp|p:cxnSp/a:spPr/a:effectLst/a:outerShdw/@rotWithShape
```

The existing strict shadow reader preserves the optional `rotWithShape`
attribute. This slice exposes it only when the attribute is present and uses
the canonical `0`/`1` boolean form; the shadow's color, geometry and all other
effect topology remain source-owned.

## Safety boundary

Missing `rotWithShape`, multiple effect lists, multiple effects, non-direct or
extension children, malformed topology, text-box/placeholder ownership and
relationship changes remain source-owned or fail closed. No source-bound
insertion, removal, or reordering is allowed.

## Edit proof

The leaf kind is `shadowRotateWithShape` and its PPJ location is
`style.shadow.rotateWithShape`. The edit plan re-proves the selected shape,
direct shape properties, bounded effect list, direct `a:outerShdw`, and
explicit `rotWithShape` attribute before token-splicing only that attribute in
the owning `SlidePart`. No protobuf or wire-version change is needed. The
focused source-bound shape-effect fixture adds the explicit authored boolean,
changes the leaf, verifies SlidePart-only mutation and package validity, and
reprojects the value.
