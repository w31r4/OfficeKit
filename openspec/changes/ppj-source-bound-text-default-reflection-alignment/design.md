# Design: source-bound paragraph default-text reflection alignment

## Owner

The owner is one direct paragraph default-run properties graph:

```text
p:sp/a:txBody/a:p/a:pPr/a:defRPr/a:effectLst/a:reflection
```

The strict reader accepts one full-span direct `a:reflection` with a bounded
`algn` token, optionally preceded by one already verified direct outer shadow.
Existing bounded reflection scalars and transforms may remain on the same
owner; the new leaf is independent of them. The paragraph index identifies the
`a:defRPr` owner and keeps it separate from inline text reflection leaves.

## Safety boundary

The owner must remain a single direct effect list and a full-span reflection.
Unknown attributes, child content, duplicate effects, `rotWithShape`, malformed
alignment tokens, and other unmodeled effect graphs remain source-owned or fail
closed. No source-bound insertion, removal, reordering, or effect normalization
is allowed.

## Edit proof

The leaf kind is `textDefaultReflectionAlignment` and its PPJ location is
`paragraph.style.defaultText.reflection.alignment`. The edit plan re-proves the
selected paragraph, direct `pPr`, `defRPr`, strict full-span effect list, and
direct reflection before token-splicing only `@algn` in the owning SlidePart.
The focused fixture removes the embedded PPJ, projects an imported `b` token,
changes it to `tr`, validates the XML and changed-part footprint, and
reprojects the new value. No protobuf or wire-version change is needed.
