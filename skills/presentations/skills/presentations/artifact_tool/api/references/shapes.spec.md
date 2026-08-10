# Shapes

`slide.shapes` creates and edits drawable slide elements.

## Preset Shape

```ts
const shape = slide.shapes.add({
  geometry,
  name,
  position,
  fill,
  line,
  accessibility,
  adjustmentList,
  borderRadius,
  shadow,
  className,
});

shape.text = textValue;
shape.text.style = {
  styleName: textStyleName,
  fontSize,
  bold,
  color,
};
shape.position = nextPosition;
```

`geometry` is a string preset shape name, `"textbox"`, `"line"`, `"connector"`, or `"custom"`.

## Resolved From Inspect

```ts
const shape = presentation.resolve("sh/a1b2c3d4");
shape.text = "Updated copy";
shape.fill = "accent1+18/90";
shape.line = { style: "solid", fill: "slate-200", width: 1 };
shape.borderRadius = "rounded-2xl";
shape.shadow = "shadow-sm";
```

Use `presentation.inspect({ kind: "shape,textbox", search })` to find the
`sh/...` anchor id. Keep the resolved facade type-aware; do not rebuild an
imported shape unless the task requires a new object.

## Non-visible accessibility metadata

Ordinary shapes may carry PowerPoint alternative-text metadata independently
of their visible text and inspectable object name:

```ts
const status = slide.shapes.add({
  name: "rollout-status",
  geometry: "roundRect",
  position: { left: 72, top: 144, width: 360, height: 120 },
  accessibility: {
    title: "Controlled rollout status",
    description: "The rollout is approved for two pilot regions.",
    decorative: false,
  },
});

if (status.accessibilityCapability.editable) {
  status.setAccessibilityMetadata({
    title: "Pilot rollout status",
    description: null,
  });
}
```

Each present string contains 1–1,024 XML-safe characters. `null` clears one
field. `decorative` is a presence-aware boolean: `false` is distinct from an
omitted classification, while `true` cannot coexist with title or description.
Switching classifications is one transaction, for example
`{ title: null, description: null, decorative: true }`. Canonical imported
`p:nvSpPr/p:cNvPr` metadata is source-bound editable; the Office 2019+
classification maps through the standard decorative extension. Hyperlinks,
unknown extensions/attributes/children, duplicate decorative extensions, and
malformed values remain byte-preserved but fail closed on semantic mutation.
This bounded metadata does not establish slide reading order or whole-deck
accessibility. Images have their separate residual-protected contract.

## Imported Background Fill

`shape.useBackgroundFill` is `true` or `false` only when imported native
PresentationML carried `p:sp/@useBgFill`; otherwise it is `undefined`. The
property is read-only. It affects preview paint, while OfficeKit preserves
the native attribute unchanged and rejects source-free authoring or mutation.

```ts
const inheritsSlideBackgroundPaint = shape.useBackgroundFill;
```

## Preset Shape Inline Type

```ts
type ShapePresetName = string; // common: "textbox", "rect", "roundRect", "ellipse", "line", "rightArrow"
// Full preset list: rg "SHAPE_GEOMETRY_NAME_TO_PROTO" src/models/presentation src/assets

type PositionConfig = {
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  rotation?: number;
  horizontalFlip?: boolean;
  verticalFlip?: boolean;
};

type PresetShapeConfig = {
  geometry: ShapePresetName;
  name?: string;
  position?: PositionConfig;
  fill?: FillConfig;
  line?: LineConfig;
  accessibility?: { title?: string; description?: string; decorative?: boolean };
  adjustmentList?: Array<{ name: string; formula: string }>;
  borderRadius?: number | string; // number = pixels; string = supported rounded-* token
  shadow?: string; // shadow token, "shadow-none", or custom "2px 7px 19px #000000/17"
  className?: string;
};
```

Placeholder authoring is layout-driven rather than a generic shape option.
For a new source-free deck, define a direct-frame `title`, `body`, `ctrTitle`,
or `subTitle` placeholder on the canonical master/layout, then call
`slide.setLayout(layout)` to materialize the slide shape. Imported `p:ph`
shapes remain inspectable but source-bound and read-only; picture/chart/table
placeholder classes and arbitrary inherited template geometry are outside this
API boundary.

## Rounded Corners

Prefer `borderRadius` for rect-like rounded corners:

```ts
slide.shapes.add({
  geometry: "rect",
  position: { left: 80, top: 120, width: 320, height: 160 },
  fill: "white",
  borderRadius: "rounded-2xl",
});

shape.borderRadius = 24;
```

Numbers are pixels. Strings use supported `rounded-*` tokens. `borderRadius`
requires a shape with width/height and is supported for `rect`, `textbox`, and
`roundRect`-like shapes. Use `adjustmentList` only when you need exact OpenXML
preset adjustment formulas.

## Rounded Rect Adjustment

```ts
const rounded = slide.shapes.add({
  geometry: "roundRect",
  position,
  fill,
  line,
  adjustmentList: [{ name: "adj", formula: adjustmentFormula }],
});
```

The `adj` value uses the OpenXML shape adjustment scale for the corner radius.

## Custom Path

```ts
const custom = slide.shapes.add({
  geometry: "custom",
  position,
  fill,
  line,
  textRectangle: { left: 24, top: 18, right: 156, bottom: 96 },
  customPaths: [
    {
      width,
      height,
      fillMode: "none",
      stroke: false,
      extrusionAllowed: false,
      commands,
    },
  ],
});
```

Custom path coordinates are signed 32-bit integer units. Positive
`width`/`height` values define an explicit path viewport that OfficeKit scales
to the shape frame. Omit either axis (or pass its native default `0`) to use the
shape's EMU coordinate system for that axis.
An arc keeps DrawingML's native radii and 1/60000-degree angles; it requires a
current point and accepts a non-zero clockwise or counter-clockwise sweep of at
most one full turn.

An ordered DrawingML formula graph can parameterize point coordinates and arc
radii/angles:

```ts
const formulaTriangle = slide.shapes.add({
  geometry: "custom",
  position: { left: 40, top: 40, width: 560, height: 280 },
  fill: "#16A34A",
  customAdjustments: [{ name: "adjX", formula: "val 25000" }],
  customGuides: [
    { name: "apexX", formula: "*/ w adjX 100000" },
    { name: "textLeft", formula: "*/ w 1 10" },
    { name: "textRight", formula: "*/ w 9 10" },
  ],
  textRectangle: { left: "textLeft", top: "t", right: "textRight", bottom: "b" },
  customConnectionSites: [
    { angle: "3cd4", x: "hc", y: "t" },
    { angle: "cd4", x: "hc", y: "b" },
  ],
  customAdjustmentHandles: [{
    kind: "xy",
    xAdjustment: "adjX",
    minX: "l",
    maxX: "r",
    x: "apexX",
    y: "vc",
  }],
  customPaths: [{
    commands: [
      { moveTo: { x: "l", y: "b" } },
      { lineTo: { x: "apexX", y: "t" } },
      { lineTo: { x: "r", y: "b" } },
      { close: {} },
    ],
  }],
});
```

`customAdjustments` is the ordered native `a:avLst` (at most 256 entries),
followed by `customGuides` as `a:gdLst` (at most 1,024 entries). Each entry is
`{ name, formula }`; names are bounded ASCII identifiers and the `officeKit`
prefix is reserved for codec-owned text-rectangle guides. Formulas are at most
256 characters and use exactly one of DrawingML's 17 operations:

```text
*/  +-  +/  ?:  abs  at2  cat2  cos  max  min  mod  pin  sat2  sin  sqrt  tan  val
```

Operands may be signed integer literals, an earlier adjustment/guide, or one
of the supported DrawingML built-ins (`3cd4`, `3cd8`, `5cd8`, `7cd8`, `b`,
`cd2`, `cd4`, `cd8`, `h`, `hc`, `hd2`, `hd3`, `hd4`, `hd5`, `hd6`, `hd8`,
`l`, `ls`, `r`, `ss`, `ssd2`, `ssd4`, `ssd6`, `ssd8`, `ssd16`, `ssd32`, `t`,
`vc`, `w`, `wd2`, `wd3`, `wd4`, `wd5`, `wd6`, `wd8`, `wd10`). Evaluation is
strictly in declaration order; forward/unknown references, duplicate names,
division by zero, negative square roots, non-finite results, and results
outside the signed 32-bit profile fail closed.

A string in a path, connection-site, handle-bound, handle-position, or text-
rectangle field may name a supported DrawingML built-in or the exact name of a
declared adjustment/guide. An explicit path `width`/`height` remains a positive
literal; omission uses the shape-coordinate default and is the most direct fit
for shape-derived `w`/`h` guides. Unknown/forward references, unsupported
formula syntax, and handle topology outside the bounded profile keep an
imported shape opaque and source-bound.

`customAdjustmentHandles` is the ordered native `a:ahLst` table, with at most
1,024 entries. An `xy` handle controls at least one declared
`xAdjustment`/`yAdjustment`; a `polar` handle controls at least one declared
`radialAdjustment`/`angleAdjustment`. Every controlled dimension either omits
its bounds or supplies its min/max pair. Coordinate and radius bounds are
signed DrawingML adjustment units, radial bounds must evaluate non-negative,
and numeric angle bounds are degrees. Bounds may instead name a built-in or
declared adjustment/guide. Numeric handle x/y positions are shape-local pixels
and may also use those references. The formula graph must place the current
adjustment inside every supplied range and the resolved handle position inside
the shape frame.

Array order is native identity. On a recognized import, an edit may change
paired bounds and position, but it cannot add/remove/reorder a handle, change
`xy` to `polar`, or retarget the controlled adjustment names. Unknown children,
attributes, missing range pairs, and broader handle topology make the whole
custom shape opaque; OfficeKit never drops them to make an edit succeed.

`customConnectionSites` is the ordered native `a:cxnLst` table, with at most
1,024 `{ angle, x, y }` entries. Numeric angles are degrees; numeric x/y values
are pixels relative to the shape frame. Any field may instead name a built-in
or declared adjustment/guide. Resolved positions must remain inside the shape
and angles within one turn. The array index is the native connector identity:
a recognized import may edit values at existing indexes but cannot change the
list length. Connectors targeting a custom shape must use explicit
`fromIdx`/`toIdx`; side aliases are intentionally limited to preset geometries.

OfficeKit's SVG/sharp render gate evaluates the formula graph. A bounded
LibreOffice/Poppler regression also proves that changing one adjustment moves
the native-rendered built-in/default-extent path in the same direction. This is
a compatibility smoke rather than a universal formula oracle; require a
Microsoft PowerPoint/native-host review when exact cross-host formula rendering
is release-critical.

`fillMode` is optional and accepts only `"normal"` or `"none"`. Omitting it
preserves an omitted native `fill` attribute (whose DrawingML default is
`norm`); `"normal"` writes an explicit `fill="norm"`, while `"none"` disables
the path fill. Optional `stroke` writes the native path-stroke boolean, whose
omitted default is true. Optional `extrusionAllowed` preserves the native
`extrusionOk` eligibility flag; it does not author a 3D scene or promise a 3D
preview. Relative `lighten`, `lightenLess`, `darken`, and `darkenLess` path
fills remain opaque because the static preview does not implement their native
paint transform.

`textRectangle` is optional and belongs to the custom shape, not to an
individual path. Each `left`, `top`, `right`, and `bottom` value is either a
pixel coordinate or a DrawingML built-in/declared adjustment/guide name. The
resolved rectangle may inset or extend the native text box, but right and
bottom must remain greater than left and top. OfficeKit retains its
deterministic four-guide scaling profile for numeric edges and writes reference
edges directly as standard `a:rect` `ST_AdjCoordinate` values; mixed rectangles
round-trip through the same model. The state drives inspect, static text
origin, and overflow QA and survives source-bound edits. Omit the field for
DrawingML's full-shape default. Unknown references, malformed leaves, or
invalid resolved bounds preserve an imported shape as opaque and reject
semantic mutation.

```ts
const ellipsePath = {
  width: 21_600,
  height: 21_600,
  commands: [
    { moveTo: { x: 10_800, y: 20_000 } },
    {
      arcTo: {
        widthRadius: 3_000,
        heightRadius: 4_000,
        startAngle: 5_400_000,
        sweepAngle: 21_600_000,
      },
    },
    { close: {} },
  ],
};
```

## Custom Path Inline Type

```ts
type CustomShapeConfig = Omit<PresetShapeConfig, "geometry"> & {
  geometry: "custom";
  customAdjustments?: Array<{ name: string; formula: string }>;
  customGuides?: Array<{ name: string; formula: string }>;
  customConnectionSites?: Array<{
    angle: number | string;
    x: number | string;
    y: number | string;
  }>;
  customAdjustmentHandles?: Array<
    | {
        kind: "xy";
        xAdjustment?: string;
        minX?: number | string;
        maxX?: number | string;
        yAdjustment?: string;
        minY?: number | string;
        maxY?: number | string;
        x: number | string;
        y: number | string;
      }
    | {
        kind: "polar";
        radialAdjustment?: string;
        minRadius?: number | string;
        maxRadius?: number | string;
        angleAdjustment?: string;
        minAngle?: number | string;
        maxAngle?: number | string;
        x: number | string;
        y: number | string;
      }
  >;
  textRectangle?: { left: number | string; top: number | string; right: number | string; bottom: number | string };
  customPaths: Array<{
    width?: number;
    height?: number;
    fillMode?: "normal" | "none";
    stroke?: boolean;
    extrusionAllowed?: boolean;
    commands: Array<
      | { moveTo: { x: number | string; y: number | string } }
      | { lineTo: { x: number | string; y: number | string } }
      | { quadraticBezTo: { x1: number | string; y1: number | string; x: number | string; y: number | string } }
      | { cubicBezTo: { x1: number | string; y1: number | string; x2: number | string; y2: number | string; x: number | string; y: number | string } }
      | { arcTo: { widthRadius: number | string; heightRadius: number | string; startAngle: number | string; sweepAngle: number | string } }
      | { close: Record<string, never> }
    >;
  }>;
};
```

## Connector

```ts
const connector = slide.shapes.connect(sourceShape, targetShape, {
  kind: "elbow",
  fromSide: "right",
  toSide: "left",
  line: { style: "solid", fill: "slate-400", width: 2 },
  head: { type: "arrow", width: "med", length: "med" },
});
```

Use [`connectors.md`](./connectors.md) when endpoints must retain target-shape
and connection-site identity, including routing, side anchors, direct
`geometry: "connector"` creation, and endpoint edits.

For `geometry: "custom"`, pass `fromIdx` or `toIdx` explicitly. The index
addresses that shape's ordered `customConnectionSites`; `fromSide`/`toSide`
exist only for the bounded preset site maps.

## Line Primitive Decision

| Need | Use |
| --- | --- |
| Divider inside compose JSX | `<rule stroke="slate-200" weight={1} />` |
| Free-positioned line | `slide.shapes.add({ geometry: "line", position, fill: "none", line })` |
| Free-positioned arrow | `slide.shapes.add({ geometry: "line", position, line: { ...line, tail } })` |
| Arrow connected to shapes | `slide.shapes.connect(fromShape, toShape, { line, head })` |
| Border around a surface | shape or box `line={{ style: "solid", fill: "slate-200", width: 1 }}` |

A free-positioned line is an ordinary shape, not a connector:

```js
const divider = slide.shapes.add({
  name: "section-divider",
  geometry: "line",
  position: { left: 72, top: 122, width: 1108, height: 0 },
  fill: "none",
  line: {
    style: "dash-dot",
    fill: "slate-500",
    width: 1.5,
    head: { type: "oval", width: "sm", length: "med" },
    tail: { type: "arrow", width: "lg", length: "sm" },
    cap: "round",
    join: "bevel",
  },
});
```

For `geometry: "line"`, `position.left/top` is the start point and
`position.width/height` is the non-negative delta to the endpoint. Horizontal
and vertical lines therefore use one zero extent; both extents zero fail
closed. OfficeKit writes this as `p:sp` with `a:prstGeom prst="line"`. It has no
target shape or connection-site index.

The bounded outline styles are `solid`, `dashed`, `dotted`, `dash-dot`,
`dash-dot-dot`, and `none`. The input aliases `dash`, `dot`, `dashDot`, and
`longDashDotDot` normalize to those canonical names. Unknown styles fail
closed. Free lines also accept `head` and `tail` with
`triangle|stealth|diamond|oval|arrow`, independent `sm|med|lg` width/length,
plus `flat|round|square` cap and `round|bevel|miter` join. The legacy flat
`startArrow*`/`endArrow*` aliases remain accepted when they do not conflict with
the nested values. Arrowheads on non-line shapes fail closed; caps and joins
may style other ordinary shape outlines. Target attachment and rerouting still
require a real `slide.shapes.connect(...)` connector.

```ts
type ShapeLineEnd = {
  type: "triangle" | "stealth" | "diamond" | "oval" | "arrow";
  width?: "sm" | "med" | "lg";
  length?: "sm" | "med" | "lg";
};

type ShapeLineConfig = LineConfig & {
  head?: ShapeLineEnd | "none" | false;
  tail?: ShapeLineEnd | "none" | false;
  cap?: "flat" | "round" | "square";
  join?: "round" | "bevel" | "miter";
};
```

Canonical free lines support authoring, import, source-bound style/frame edits,
slide duplication, export, and second import. Imported theme colors, custom
dash graphs, compound lines, effects, extension content, missing/ambiguous line
fills, or otherwise complex `a:ln` content stay source-bound: unchanged export
preserves the package, while semantic mutation fails closed instead of
flattening the outline.

## Shadows

Supported shadow tokens:

```text
shadow-none, shadow-sm, shadow, shadow-md, shadow-lg, shadow-xl, shadow-2xl
```

Custom shadow strings are also supported when the presentation has a theme
context:

```ts
shape.shadow = "shadow-md";
shape.shadow = "2px 7px 19px #000000/17";
shape.shadow = "shadow-none";
```

## Ordering

```ts
shape.bringToFront();
shape.sendToBack();
```

## Cookbook

```ts
// KPI metric surface.
const metricSurface = slide.shapes.add({
  geometry: "roundRect",
  name: "kpi-surface",
  position: { left: 64, top: 132, width: 260, height: 148 },
  fill: "white",
  line: { style: "solid", fill: "slate-200", width: 1 },
  borderRadius: "rounded-2xl",
  shadow: "shadow-md",
});
metricSurface.text = "Revenue\n$12.4M";
metricSurface.text.style = {
  className: "text-slate-950 text-2xl font-bold leading-tight",
};
```

```ts
// Pill label.
const pill = slide.shapes.add({
  geometry: "roundRect",
  position: { left: 64, top: 64, width: 148, height: 34 },
  fill: "emerald-50",
  line: { style: "solid", fill: "emerald-200", width: 1 },
  borderRadius: "rounded-full",
});
pill.text = "ON TRACK";
pill.text.style = { className: "text-emerald-800 text-sm font-bold" };
```

```ts
// Directional connector between two shapes.
slide.shapes.connect(sourceShape, targetShape, {
  kind: "elbow",
  fromSide: "right",
  toSide: "left",
  line: { style: "solid", fill: "slate-400", width: 2 },
  head: { type: "triangle", width: "sm", length: "sm" },
});
```
