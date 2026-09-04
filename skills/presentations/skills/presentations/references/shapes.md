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
`cubicTo`, `arcTo`, and `close` commands. Path order is drawing order; `fill`
and `stroke` control whether the shared shape paint applies to that path. A
non-zero view-box origin is normalized deterministically during compilation.
`arcTo` uses positive `radiusX`/`radiusY` values in view-box units plus
`startAngle` and signed `sweepAngle` in degrees. It requires a current point;
start a contour with `moveTo`. Positive sweep follows the clockwise page
coordinate convention. Use opposite sweeps for the outer and inner contours of
a native hollow ring. The same command vocabulary applies to custom image
masks.
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

An imported literal custom geometry (only paths with literal coordinates, no
guides, adjustment handles, connection sites, or text rectangle) may likewise
issue `setGeometry` for `geometry.paths`. That bounded edit replaces the
shape-owned path list in the existing SlidePart and reprojects it; formula or
extension-bearing custom geometry remains source-bound and is not flattened.

The same preset profile can clip an image. `image.mask.adjustments` uses the
identical parameter order and defaults; see [Media and layers](media-and-layers.md#image-masks).
Connector presets are intentionally absent from shape geometry because PPJ has
a typed connector element with endpoint semantics.

## Named icons are native vectors

Use `type: "icon"` for a small, familiar symbol whose meaning is clearer than
a label or decorative shape. `iconName` selects one exact bundled Font Awesome
Free 7.3.1 name using `fas:`, `far:`, or `fab:`. The compiler fits its original
aspect ratio inside the declared frame and writes one editable DrawingML custom
shape. It does not introduce an image asset, relationship, remote request, font
dependency, or SVG runtime.

```json
{
  "type": "icon",
  "id": "insight-symbol",
  "iconName": "fas:lightbulb",
  "frame": { "x": 824, "y": 48, "width": 40, "height": 40 },
  "style": {
    "fill": { "type": "solid", "color": { "token": "signal" } }
  },
  "accessibility": {
    "description": "A lightbulb marks the central experimental insight."
  }
}
```

Choose an icon only when its conventional meaning is unambiguous to the
audience. Prefer text for uncommon concepts, data graphics for evidence, and
an image for identity or atmosphere. Do not repeat icons as page filler or use
them as bullet decoration. Brand icons identify their owner, product, or
service only; they do not imply endorsement. Mark a purely redundant icon as
decorative, otherwise provide a short accessible description.

The catalog contains 2,163 pinned names. Search the bundled
`src/ppj/font-awesome-free-icons.json` when the exact spelling is uncertain;
`ppj check` rejects an unknown name rather than substituting another symbol.
An ordinary imported custom shape remains a shape or opaque native object: the
projector never guesses an `iconName` from geometry. Exact embedded PPJ recovery
does retain the original semantic icon element.

## Lines are connectors

Use `connector` as PPJ's ordinary line primitive. A connector may join two
literal points or bind to stable element IDs, and it owns straight, elbow, or
curved routing plus stroke and arrowheads. Do not use custom geometry merely to
draw a line, and do not introduce a second line element with overlapping
semantics.

```json
{
  "type": "connector",
  "id": "threshold-line",
  "connectorType": "straight",
  "from": { "x": 84, "y": 260 },
  "to": { "x": 520, "y": 260 },
  "stroke": { "color": "#16324F", "width": 1.5, "dash": "dash" },
  "endArrow": "none"
}
```

Bind endpoints to elements when the relationship must survive movement. Use
literal points for rules, axes, baselines, thresholds, or deliberate visual
dividers that do not belong to another object.

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

This native half-ellipse starts at the current point; OfficeKit writes an
editable DrawingML arc rather than approximating it with a bitmap or requiring
the Agent to calculate Bézier control points:

```json
{
  "geometry": {
    "kind": "custom",
    "viewBox": { "x": 0, "y": 0, "width": 100, "height": 100 },
    "paths": [{
      "fill": false,
      "stroke": true,
      "commands": [
        { "op": "moveTo", "x": 0, "y": 50 },
        { "op": "arcTo", "radiusX": 50, "radiusY": 50, "startAngle": 180, "sweepAngle": 180 }
      ]
    }]
  }
}
```

Stroke opacity is native DrawingML alpha, not a flattened visual effect. It is
available on authored shapes and connectors and survives import, PPJ projection,
source-bound edits, and rebuild. Use it to keep a secondary relationship
present without competing with the evidence carrier; do not make required axes
or data lines faint.

Prefer alpha on the branch that actually needs it: fill, stroke, image, border,
shadow, or gradient stop. Use `shape.style.opacity` when the entire authored
shape must fade as one semantic object. OfficeKit multiplies that value into
each directly owned solid, gradient or image fill, outline, shadow, explicitly
painted text, text shadow and bullet color, preserving any branch-local alpha.
The result remains one native editable shape. Inherited text paint and text
highlight fail closed under compound opacity because resolving them would
invent a color or an unsupported highlight alpha.

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

Recognized imported ordinary shapes and lines may expose one direct outer
shadow as `style.shadow` (or `shadow` on a line). A `setShapeEffects`
capability permits changing or clearing its RGB/theme color, blur, distance,
angle, alignment, rotation behavior, and opacity while retaining the existing
geometry, paint, text, and native relationships. Text boxes, placeholders,
multi-effect/extension graphs, glow, reflection, inner shadow, soft edge, and
3-D effects remain source-owned and fail closed rather than being flattened.

## Authored semantic diagrams

Use `type: "smartArt"` with `mode: "authored"` when the content is genuinely a
finite list, process, cycle, hierarchy, relationship, matrix, pyramid, or
picture sequence. OfficeKit compiles one element to one native SmartArt
graphic frame with data, layout, style, colors, and a deterministic cached
drawing. The cached drawing is internal to the SmartArt object; its shapes are
not exposed as independent page elements. Imported native SmartArt uses
`mode: "source-bound"` and remains limited to its issued `nativeRef`
capabilities. When the DiagramML graph is fully proven, PPJ may expose its
immutable `layoutDefinitionId`, content-node `kind`, and `parent` connections;
these are an inspectable semantic facade, not authority to rebuild or reparent
the source graph.

The program must supply named shape and text styles. Connected layouts also
supply connector paint. This keeps the compiler deterministic without letting
it invent a palette, typography system, or decorative geometry:

```json
{
  "id": "evidence-chain",
  "type": "smartArt",
  "frame": { "x": 72, "y": 160, "width": 816, "height": 150 },
  "mode": "authored",
  "layout": "process",
  "shapeStyleRef": "evidence-stage",
  "textStyleRef": "stage-label",
  "nodeGeometry": { "kind": "preset", "preset": "roundRect" },
  "connector": {
    "stroke": { "color": { "token": "signal" }, "width": 1.5 },
    "endArrow": "triangle"
  },
  "nodes": [
    { "id": "observe", "text": "Observe" },
    { "id": "measure", "text": "Measure" },
    { "id": "decide", "text": "Decide" }
  ],
  "connections": [
    { "id": "observe-measure", "from": "observe", "to": "measure", "role": "sequence", "order": 0 },
    { "id": "measure-decide", "from": "measure", "to": "decide", "role": "sequence", "order": 1 }
  ]
}
```

Use `connections` as the only topology language: `sequence` for process/cycle,
`parent` for hierarchy, and `association` for relationship graphs. Ordered
nodes still determine stable placement for list, matrix, pyramid, and picture.
Picture nodes each declare an image `asset`. A node may override
`shapeStyleRef`, `styleRef`, or `geometry`.
The authored budget is 1–64 nodes. For a composition whose layout itself is the
message, use an explicit `group` and frames instead of forcing it into one of
these eight bounded layouts.

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
