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

`geometry` is a string preset shape name, `"textbox"`, `"connector"`, or `"custom"`.

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

Custom path coordinates are signed 32-bit integer units in the path's own
`width`/`height` viewport. OfficeKit scales that viewport to the shape frame.
An arc keeps DrawingML's native radii and 1/60000-degree angles; it requires a
current point and accepts a non-zero clockwise or counter-clockwise sweep of at
most one full turn. Formula-valued native geometry remains opaque.

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
individual path. Its `left`, `top`, `right`, and `bottom` values are pixels
relative to the shape frame; the rectangle may inset or extend the native text
box. OfficeKit reads native numeric rectangles and writes one deterministic
four-guide DrawingML `a:rect` profile so PowerPoint and LibreOffice resolve the
same shape-local EMUs. The same state drives inspect, static text origin, and
overflow QA and survives source-bound edits. Omit the field for DrawingML's
full-shape default. Every other formula-valued native rectangle remains opaque
and rejects semantic mutation rather than exposing a partial formula model.

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
  textRectangle?: { left: number; top: number; right: number; bottom: number };
  customPaths: Array<{
    width: number;
    height: number;
    fillMode?: "normal" | "none";
    stroke?: boolean;
    extrusionAllowed?: boolean;
    commands: Array<
      | { moveTo: { x: number; y: number } }
      | { lineTo: { x: number; y: number } }
      | { quadraticBezTo: { x1: number; y1: number; x: number; y: number } }
      | { cubicBezTo: { x1: number; y1: number; x2: number; y2: number; x: number; y: number } }
      | { arcTo: { widthRadius: number; heightRadius: number; startAngle: number; sweepAngle: number } }
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

Use [`connectors.md`](./connectors.md) for connector routing, side anchors,
direct `geometry: "connector"` creation, arrowheads, and endpoint edits.

## Line Primitive Decision

| Need | Use |
| --- | --- |
| Divider inside compose JSX | `<rule stroke="slate-200" weight={1} />` |
| Free-positioned line | `slide.shapes.add({ geometry: "line", position, fill: "none", line })` |
| Arrow connected to shapes | `slide.shapes.connect(fromShape, toShape, { line, head })` |
| Border around a surface | shape or box `line={{ style: "solid", fill: "slate-200", width: 1 }}` |

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
