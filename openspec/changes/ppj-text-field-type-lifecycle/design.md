# Design

## Source-bound field

`text.paragraphs[].runs[].field.type` is the public field. A fresh source
projection issues a `textFieldType` native leaf for each non-automatic field
whose type satisfies the existing 1..255 printable field-type grammar. The
leaf index is the field ordinal in the shape, including automatic fields that
do not receive a leaf, so the edit plan can re-prove the exact `a:fld` node.

The element advertises `setTextField` with the field path
`text.paragraphs[].runs[].field.type` when at least one such leaf exists.
Compilation requires that capability and rejects changes to the field ID,
cached text, automatic marker, run/paragraph topology, or any automatic field
type. Both the old and new type must remain static and valid.

## Token splice

The edit plan validates the source slide, shape binding, field ordinal,
expected source type, and static-field profile. It replaces only the value of
the direct `type` attribute on that `a:fld` start tag, XML-escaping the new
token. The original package is untouched; the output scope is the owning slide
part. Reprojection reads the new type while preserving the field ID and cached
display text.

## Evidence boundary

Focused tests cover authored and imported static fields, capability issuance,
one-slide byte scope, Open XML validation, reprojected type, and rejection of
automatic-field or identity changes. This does not claim host PowerPoint
field refresh or evaluation support.
