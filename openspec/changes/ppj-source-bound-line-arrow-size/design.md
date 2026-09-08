# Design: source-bound line arrow size

## Owner and projection

`lineStartArrowWidth` and `lineStartArrowLength` bind to a direct shape or
connector `a:ln/a:headEnd/@w` and `@len`. `lineEndArrowWidth` and
`lineEndArrowLength` bind to the corresponding `a:tailEnd` attributes.

The projector issues a leaf only for an existing explicit `headEnd` or
`tailEnd` with one canonical arrow type and an explicit `sm|med|lg` width or
length token. The owner must have the existing bounded direct line profile;
unknown children, extensions, duplicate endpoint owners, and unsupported
arrow types remain source-owned.

## Edit boundary

The source-bound edit plan accepts only the four canonical size tokens and
only when the expected value matches the hash-bound source. The XML writer
replaces exactly the selected endpoint attribute in the owning SlidePart.
Arrow type, opposite endpoint, line paint, dash/cap/join, geometry, effects,
and all non-owner parts remain untouched. A second projection must recover the
new leaf value.

No new protobuf field is required: the native model already carries the four
arrow size strings.
