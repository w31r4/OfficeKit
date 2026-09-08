# Design

## Boundary

The source-bound compound opacity profile applies only to an ordinary,
non-placeholder `shape` with no visible text. An empty native text container
is harmless and remains untouched.
The projector collects effective alpha from each modeled visual owner:

- direct solid fill;
- every stop in a bounded gradient;
- a canonical image fill;
- a visible colored outline; and
- an imported shadow.

An omitted alpha means `1`. The field is issued only when at least one owner
exists and all collected values are equal within the native thousandth-percent
rounding tolerance. A source-owned custom image-fill identity without a
canonical `PresentationImagePaint` is not included in this profile.

## Lowering and recovery

`compositing.opacity` is lowered as an absolute value by setting or clearing
the alpha on every modeled owner, rather than multiplying an already projected
value. This makes repeated source-bound edits deterministic and avoids a
special case around zero. Existing DrawingML fill, gradient, outline, image,
and shadow codecs remain the native owners.

The focused experiment projects a source-bound shape, edits only
`compositing.opacity`, and verifies that the source package changes only in
the owning slide XML and that both the fill and outline carry the requested
alpha. A subsequent projection recovers the same compound opacity.

## Source-bound edit rules

- The element capability must contain `setOpacity` for
  `compositing.opacity`.
- A transaction may not change `style` and `compositing` together; otherwise
  local paint edits could alter the owner set while opacity is being applied.
- `blendMode`, `isolation`, and `clipStack` remain source-owned/unsupported.
- Lines keep their direct `stroke.opacity` owner; text-bearing shapes keep
  their existing text and direct paint semantics; placeholders remain on the
  placeholder path.

## Non-goals

- Group, page, or layer transparency and compositing order.
- Non-normal blend modes, isolation, or arbitrary clip stacks.
- Text alpha, effect graphs beyond the modeled shadow, or opaque custom
  image-fill topology.
- Adding a new protobuf field or changing the Office wire version.
