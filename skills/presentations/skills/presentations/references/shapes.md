# Shapes and diagrams

Decide the page's core message before drawing. A shape earns space only when it
encodes data, a relationship, a boundary, focus, or identity.

## Choose geometry by meaning

- Use position, scale, alignment, and shared baselines for comparison.
- Use lines and connectors for direction, dependency, sequence, or causality.
- Use bounded regions for real states, phases, ownership, or physical areas.
- Use native vectors for a deliberate identity motif when they carry the same
  meaning across the deck.
- Use a chart or table instead of a diagram when quantity or exact values are
  the relationship.

PPJ shape elements use a stable `id`, explicit `frame`, typed `geometry`, and
an optional named style. Connector endpoints should bind stable element IDs
when the relationship belongs to those objects. Keep arrow direction and label
placement unambiguous.

Shape, image, chart, table, and group frames support explicit `rotation`,
`flipH`, and `flipV`. A group transform changes the outer group frame while its
children keep their owner-local coordinates. Connector orientation is derived
from its endpoints, so do not add an independent connector rotation.

Custom geometry is justified only when a preset cannot express a necessary
semantic form. Keep its path finite and editable. Do not trace a decorative
illustration into arbitrary geometry when an image or SVG asset is the honest
carrier.

PPJ compiles literal custom paths directly into editable DrawingML geometry.
Use one finite `viewBox` plus ordered `moveTo`, `lineTo`, `quadraticTo`,
`cubicTo`, and `close` commands. Path order is drawing order; `fill` and
`stroke` control whether the shared shape paint applies to that path. A
non-zero view-box origin is normalized deterministically during compilation.
Custom adjustment formulas, guides, handles, connection sites, and text
rectangles remain outside the authored PPJ subset. Preset geometry is broader
and covers the complete non-connector DrawingML catalog owned by the pinned
Office schema:
use its ordered integer `adjustments` array to control rounded corners, arrow
proportions, star radii, arc angles, callout tips, and the other parameters in
the generated [PPJ preset profile table](ppj.md#preset-geometry-adjustments).
Omit the array to use Office defaults; never invent native guide names.

```json
{
  "type": "shape",
  "id": "decision-arrow",
  "frame": { "x": 84, "y": 210, "width": 180, "height": 64 },
  "geometry": {
    "kind": "preset",
    "preset": "rightArrow",
    "adjustments": [42000, 36000]
  },
  "style": {
    "fill": { "type": "solid", "color": "#0B8F8F" }
  }
}
```

The array is complete or absent; do not omit an intermediate value. Search the
generated table instead of guessing a preset or adjustment arity. Imported
literal preset adjustments can be changed only when `nativeRef.capabilities`
issues `setGeometry` for `geometry.adjustments`. Formula-valued or irregular
native guides remain source-owned.

The same preset profile can clip an image. `image.mask.adjustments` uses the
identical parameter order and defaults; see [Media and layers](media-and-layers.md#image-masks).
Connector presets are intentionally absent from shape geometry because PPJ has
a typed connector element with endpoint semantics.

```json
{
  "type": "shape",
  "id": "signal-wave",
  "frame": { "x": 72, "y": 180, "width": 420, "height": 96 },
  "geometry": {
    "kind": "custom",
    "viewBox": { "x": 10, "y": 20, "width": 100, "height": 40 },
    "paths": [{
      "fill": false,
      "stroke": true,
      "commands": [
        { "op": "moveTo", "x": 10, "y": 40 },
        { "op": "cubicTo", "x1": 35, "y1": 10, "x2": 75, "y2": 70, "x": 110, "y": 40 }
      ]
    }]
  },
  "style": {
    "stroke": {
      "color": "#0B8F8F",
      "width": 2,
      "opacity": 0.72,
      "dash": "solid",
      "cap": "round",
      "join": "round"
    }
  }
}
```

Stroke opacity is native DrawingML alpha, not a flattened visual effect. It is
available on authored shapes and connectors and survives import, PPJ projection,
source-bound edits, and rebuild. Use it to keep a secondary relationship
present without competing with the evidence carrier; do not make required axes
or data lines faint.

Prefer alpha on the branch that actually needs it: fill, stroke, image, border,
shadow, or gradient stop. `shape.style.opacity` below one is compiler-owned only
for a solid-fill-only shape; a compound shape with text, stroke, or shadow fails
closed because a single DrawingML value cannot represent honest whole-object
opacity. Text color alpha is likewise rejected until it has a native run owner.

Use a bounded gradient only when direction or depth carries meaning. PPJ owns
linear gradients with an explicit angle and centered radial gradients with
ordered RGB stops; each stop may carry opacity. These remain editable native
DrawingML fills and survive projection and source-bound shape edits. Prefer two
or three deliberate stops. A many-color gradient used only to make a page look
busy is the same failure as random decoration.

A shape may also own a native image fill with `style.fill.type: "image"`.
`stretch`, `cover`, `contain`, explicit crop, opacity, and default tile use the
same bounded profile as a native image background. The shape geometry remains
editable and clips the picture without flattening it. Use this for a meaningful
image window or material surface, not to texture every box. See
[Media and layers](media-and-layers.md#layer-stack) for the full contract and
source-bound `setFill` rule.

## Protect reading and evidence

Order `pages[].elements[]` from back to front. Keep evidence-bearing lines,
markers, labels, numbers, axes, intervals, and sources above fills or clear of
them. A foreground shape may overlap a background field; two evidence objects
must remain traceable.

When a collision occurs, repair the composition: adjust the frame, anchor the
label, reduce an honest fill's opacity, use a local mask behind text, or change
the carrier. Do not falsify scale or separate truly related series merely to
silence an overlap check.

## Strictly forbidden

- card walls or equal rounded panels used as default hierarchy;
- colored side-strip cards, pills, badges, and button-like labels as filler;
- random circles, rings, arrows, blobs, or nodes added to make a page "rich";
- decorative process diagrams with no process relationship;
- connectors that cross labels, values, or unrelated objects;
- large empty containers whose border does all the organizing;
- a universal `box`, `card`, or `metricPanel` component driving every page.

User-provided card-based brand systems and imported layouts may be preserved.
New shapes inside them still need a declared role and clear reading order.
