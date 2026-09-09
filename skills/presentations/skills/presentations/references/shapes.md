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

Custom `geometry.paths[]` accepts optional boolean `extrusionOk` beside `fill`
and `stroke`. True, false and omission remain distinct through export and fresh
projection. Existing path-edit authority supports adding, changing or removing
it, and coordinate edits retain it. This native eligibility flag supplies no
3-D depth/material; preview reports the unrendered field explicitly.

An imported recognized custom geometry may issue `setGeometry`
for `geometry.paths`, `geometry.textRectangle`, `geometry.guides` and
`geometry.adjustments`, `geometry.connectionSites` and `geometry.adjustmentHandles`. These fields can be edited
independently in the existing SlidePart; unsupported or extension-bearing custom
geometry remains source-owned.

`geometry.textRectangle` sets the shape text area. For example:
`{ "left": 10, "top": 5, "right": "r", "bottom": "b" }`.
Numbers are points relative to the shape frame, independent of the path's
`viewBox`; strings retain native built-in references such as `l/t/r/b` or declared guide names.
All four edges are required and their resolved right/bottom must exceed
left/top. A fresh projection preserves mixed numeric/reference edges.
Source `setGeometry` permits adding, changing or removing the rectangle;
omission restores the native default. Paths, text and frame stay intact.
This is shape state; image masks and compositing clips reject it. Preview
retains explicit text-layout limitations, so this is not host layout proof.

`geometry.guides` declares ordered native formulas, for example:

```json
"guides": [
  { "name": "inset", "formula": "*/ w 1 10" },
  { "name": "rightBound", "formula": "+- w 0 inset" }
],
"textRectangle": { "left": "inset", "top": 5, "right": "rightBound", "bottom": "b" }
```

Formulas use DrawingML operators and units; `w`/`h` resolve to the shape's
native extents. Names must be unique, cannot shadow built-ins or the reserved
`officeKit` prefix, and can reference declared adjustments or earlier guides. The list is bounded
to 1024 guides; native validation rejects invalid arithmetic or unresolved
references. Source `setGeometry` can edit the list, add entries or remove it;
remove dependent rectangle references in the same request. Empty/omitted lists
project as absence; formula whitespace may normalize. Native private numeric
rectangle scaling guides are not user guides. These shape guides
are rejected on masks/clips, and preview retains explicit limitations.

For `kind: "custom"`, `geometry.adjustments` uses the same `{name, formula}`
objects as `guides`, with at most 256 entries. Adjustments are evaluated first;
each can reference built-ins or earlier adjustments, and ordinary guides can
reference all declared adjustments. For example, an adjustment
`{ "name": "padding", "formula": "val 127000" }` supplies one native EMU
coordinate value (10pt) that a later guide may reference. Names share one
namespace across both lists. Source `setGeometry` supports addition, changes
and coordinated removal; retained references must still resolve. Empty lists
project as omission. `kind: "preset"` keeps its integer adjustment array.

Custom shapes also accept `geometry.connectionSites`, an ordered list of up to
1024 `{angle, x, y}` records. Numbers are degrees and shape-local points (not
path viewBox coordinates); strings retain built-in/adjustment/guide references.
For example, `{ "angle": "cd4", "x": "hc", "y": 10 }` uses the native quarter
turn and horizontal center with a literal vertical position. Resolved positions
must lie inside the shape and angles within one signed turn. Empty authored
lists project as omission. Source `setGeometry` can change values at existing
indexes; adding/removing sites rejects because native connectors use indexes as
identity. Masks/clips reject this shape graph. Preview reports unrendered site
semantics explicitly; host snapping/dragging is unverified.

Custom `geometry.adjustmentHandles` contains up to 1024 ordered controls:

- `kind: "xy"`: `xAdjustment` with optional paired `minX/maxX`, and/or
  `yAdjustment` with optional paired `minY/maxY`.
- `kind: "polar"`: `radialAdjustment` with optional paired `minRadius/maxRadius`,
  and/or `angleAdjustment` with optional paired `minAngle/maxAngle`.

Both require `position: {x, y}` and at least one declared adjustment name.
Numbers use shape-local points for coordinates/radii and degrees for angles;
strings retain built-in/adjustment/guide references. Current adjustments must
lie inside their bounds, radii must be nonnegative and positions inside the
shape. Omitted bounds stay absent. For example:

```json
{"kind":"xy","xAdjustment":"inset","minX":0,"maxX":"w","position":{"x":"inset","y":"vc"}}
```

Source `setGeometry` can add/change/remove paired bounds and change positions;
handle order, kind and controlled names stay fixed. Empty authored lists project
as omission. Masks/clips reject controls and preview reports handle limitations;
PowerPoint dragging behavior remains unverified.

Custom-shape path commands accept a number or a native reference string in every
point/control coordinate (`x/y/x1/y1/x2/y2`) and arc parameter
(`radiusX/radiusY/startAngle/sweepAngle`). For example:

```json
{"op":"lineTo","x":"edge","y":80}
```

Numeric points retain viewBox-origin subtraction and 1000 native path units per
viewBox unit; arc angles use degrees. Strings copy native built-in/adjustment/
guide names without conversion or origin subtraction. A guide `val 20000`
therefore supplies 20000 native path units (20 viewBox units), while `w` resolves
to the native shape extent. Authored export and source `geometry.paths` edits
retain these references; changing an adjustment updates its dependent values.
Unknown references, nonpositive arc radii and invalid resolved sweeps reject.
Per-path viewports can retain different/default coordinate extents. Standalone
lines and masks/clips keep literal paths. Preview diagnoses unrendered reference
semantics explicitly.

Custom-shape paths accept `viewport: {width, height}` to override the common
`geometry.viewBox` extents. Both values are required, each from 0 to 2147483.647
in existing path units (1000 native units per unit). Omission inherits the
common extents; zero independently selects the native shape-coordinate default
on that axis. `{ "width": 0, "height": 50 }` therefore retains a default X axis
and an explicit Y extent. Command origin and reference units stay unchanged.

Source `geometry.paths` edits can add, change or remove the override; removing
it restores current viewBox inheritance. Projection factors a valid positive
baseline into geometry.viewBox and emits overrides where needed, so redundant
values normalize away. Effective native extents, commands and flags survive;
XML zero/omitted attributes share the native default meaning. Masks/clips reject
viewport overrides. Preview reports unsupported per-path viewport semantics.

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

## Connector endpoints

Use `connector` for relationships between objects or explicit endpoint pairs.
Use the independent `line` element for free paths. A connector owns straight,
elbow or curved routing, stroke and arrowheads.

```json
{
  "type": "connector",
  "id": "relation",
  "frame": { "x": 0, "y": 0, "width": 1, "height": 1 },
  "connectorType": "straight",
  "from": { "element": "source", "anchor": "right" },
  "to": { "element": "target", "anchor": "left" },
  "stroke": { "color": "#16324F", "width": 1.5 },
  "endArrow": "triangle"
}
```

`source` and `target` must identify expanded objects on the same page. Each
endpoint can instead be a literal point, for example `{ "x": 520, "y": 260 }`.
The required connector frame does not replace endpoint coordinates.

`top`, `right`, `bottom` and `left` select the target frame's side midpoint;
`center` selects its center. The compiler then applies target rotation/flips
and ancestor group transforms. `auto` chooses the shortest pair among the four
side midpoints in slide space, holding explicit endpoints fixed. Ties use
start then end order: top, right, bottom, left. Center is explicit only.

Literal coordinates belong to the connector's parent space. Component placement
transforms its literal endpoints and sibling frames together; group descendants
remain in their `childFrame` space. Endpoint coordinates are signed: negative
values locate an endpoint before the page or group origin, without clamping.
Native coordinates and extents are checked against DrawingML bounds; fresh PPJ
projection also requires its existing frame limits. Missing IDs, opaque or
connector targets, unresolved component instance ports and unrepresentable
coordinates fail with a diagnostic.

On a fresh source projection, `setConnectorType` permits changing the required
`connectorType` among `straight`, `elbow` and `curved`. The codec replaces its
recognized native geometry family while retaining endpoints, bindings, arrows
and line style.

For `elbow` and `curved`, optional `bendAdjustment` exposes the native `adj1`
literal: for example `connectorType: "elbow", bendAdjustment: 25000`. These are
DrawingML adjustment units (default midpoint 50000), not points. Signed 32-bit
integers, explicit zero and omission survive fresh projection. With
`setConnectorType`, edit the value or remove the property to restore absence.
Switching to `straight` clears an unchanged projected bend; supplying a changed
bend with straight is rejected. Endpoints, bindings and arrow state remain intact.
Formula guide graphs and nonstraight rotations outside 0/180-degree equivalents
remain opaque because endpoint normalization cannot preserve their bend axis.
This is manual adjustment, not obstacle avoidance. Nondefault internal preview
routes remain unavailable with `preview.scene.paint.connector-bend`; production
preview remains partial.

`startArrow` and `endArrow` use `none`, `triangle`, `stealth`, `diamond`,
`oval` or `open`. `open` maps to the native arrow shape and returns as `open`
on fresh projection. Editable source connectors issue `setConnectorArrows`:
change or add either field directly, or remove it/set `none` to delete that end.
A changed arrow retains its native width/length; deletion removes its own size
state. The other arrow and endpoint bindings stay intact.

Set `startArrowWidth`, `startArrowLength`, `endArrowWidth` and `endArrowLength`
independently to `sm`, `med` or `lg` (relative native sizes, not points). For
example, `endArrow: "triangle", endArrowWidth: "lg", endArrowLength: "sm"`
creates a wide, short end arrow. Omitted dimensions retain native defaults and
stay absent on fresh projection. With `setConnectorArrows`, changing one size
preserves the others; removing that property removes only its native attribute.
An authored size requires its arrow. Removing an arrow clears its unchanged
projected dimensions; explicitly changing a nonempty size while removing that
arrow is rejected. Production preview remains partial (`preview.connector.limited`);
internal scene arrows approximate contours and do not prove host-exact sizes.

On a fresh source projection, use the issued `setConnectorEndpoints` capability
to edit `from`/`to`. Moving a supported target frame or its group frame/childFrame,
including issued native frame leaves, recomputes attached endpoints. Replace an
object endpoint with a literal point to remove that attachment. Edit endpoints
or target placement rather than an attached connector's derived frame.

OfficeKit preserves the exact frame anchor in native metadata even when the
embedded PPJ snapshot is absent. This supports OfficeKit recompile/edit cycles;
PowerPoint drag attachment has not been verified. Native preset/opaque connection-site
bindings keep their separate source-owned authority. Production preview remains partial; the internal
compiler scene carries the resolved coordinates. These source-library tests do
not certify an installed NativeAOT package or host rendering.

For a supported custom shape, use `from: {element: "shape-id", connectionSite: 0}`
or the same object in `to`. The index is zero-based (0..1023) into the target's
`geometry.connectionSites`; it is exclusive with `anchor` and literal coordinates.
The compiler resolves local site coordinates/formulas through rotation, flips and
group child spaces and writes the native target/index binding. Fresh projection
retains this index and issues `setConnectorEndpoints` for supported bindings.
Change the index/target, replace the endpoint with `{x,y}` to detach, or use
`{element,anchor}` to switch to a frame anchor. Semantic target geometry/frame and
supported frame-leaf edits recompute coordinates. Geometry native-leaf edits on
a bound target reject; edit its semantic geometry instead. Missing/unsupported
targets and out-of-range indexes reject. Preview reports site-binding limitations.

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
