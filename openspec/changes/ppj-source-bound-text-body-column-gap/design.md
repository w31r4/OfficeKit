# Design: source-bound text-body column gap

## Owner

The owner is one direct text shape body property:

```text
p:sp/p:txBody/a:bodyPr/@spcCol
```

The existing text-body reader already maps this attribute to
`textBoxStyle.columnGap` in EMUs. This change only issues a separate native
leaf when the attribute is explicitly present and is a non-negative signed
32-bit integer.

## Edit proof

The leaf kind is `textBodyColumnGapEmu` and its PPJ location is
`textBoxStyle.columnGap`. The edit plan re-proves a text-editable `p:sp`, its
direct `p:txBody/a:bodyPr`, and the single existing `spcCol` attribute before
splicing only that attribute value. The value stays in native EMUs; no
protobuf or wire-version change is needed.

The focused fixture authors one text shape with a non-zero column gap,
removes the embedded PPJ snapshot, projects the native leaf, changes the
leaf, and verifies the target SlidePart, Open XML validity, preservation of
text and unrelated body attributes, and a second projection of the new gap.
