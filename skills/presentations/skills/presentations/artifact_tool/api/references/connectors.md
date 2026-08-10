# Connectors

`slide.shapes.connect` creates connector lines whose endpoints retain both the
target shape ID and the DrawingML connection-site index. Use it for arrows,
flow links, callouts, dependency lines, and other relationships between
modeled positioned shapes in one slide or group shape tree.

Use direct `geometry: "connector"` creation only when you need exact connection
site indexes.

## Connect Shapes

```ts
const connector = slide.shapes.connect(sourceShape, targetShape, {
  kind: "elbow",
  fromSide: "right",
  toSide: "left",
  line: { style: "solid", fill: "slate-500", width: 2 },
  head: { type: "arrow", width: "med", length: "med" },
  accessibility: {
    title: "Source to target flow",
    description: "Connector from the source card to the target card.",
  },
});
```

Endpoints accept shape facades or shape ids. `fromSide` and `toSide` accept
`"top"`, `"left"`, `"bottom"`, or `"right"` and resolve to the nearest
connection site for that shape geometry. When you omit side and index options,
the API chooses a connection pair from the relative shape positions.

The current side map is deliberately bounded:

| Geometry | top | left | bottom | right | Exact accepted indexes |
| --- | ---: | ---: | ---: | ---: | --- |
| `rect`, `roundRect`, `textbox` | 0 | 1 | 2 | 3 | 0-3 |
| `ellipse` | 0 | 2 | 4 | 6 | 0-7 |

Other geometries do not receive a guessed site map. `getConnectionSiteIndex`,
new attachment, or rerouting against one of them fails closed.

`kind` defaults to `"elbow"` for `slide.shapes.connect(...)`.

## Connect Inline Types

```ts
type ConnectorSide = "top" | "left" | "bottom" | "right";

type ShapeConnectOptions = {
  kind?:
    | "straight"
    | "elbow"
    | "elbow2"
    | "elbow3"
    | "elbow4"
    | "elbow5"
    | "curved";
  fromSide?: ConnectorSide;
  toSide?: ConnectorSide;
  fromIdx?: number;
  toIdx?: number;
  line?: LineConfig;
  head?: LineEndConfig;
  tail?: LineEndConfig;
  cap?: "flat" | "round" | "square";
  join?: "round" | "bevel" | "miter";
  accessibility?: { title?: string; description?: string; decorative?: boolean };
};

type LineEndConfig = {
  type?: "none" | "triangle" | "stealth" | "diamond" | "oval" | "arrow";
  width?: "sm" | "med" | "lg";
  length?: "sm" | "med" | "lg";
};
```

## Anchor Choice

Use side anchors for readable authoring:

```ts
slide.shapes.connect(sourceShape, targetShape, {
  fromSide: "bottom",
  toSide: "top",
  kind: "elbow",
  line: { style: "dashed", fill: "accent1", width: 2 },
});
```

Use explicit connection site indexes when you need an exact preset-geometry site:

```ts
const fromIdx = slide.shapes.getConnectionSiteIndex(sourceShape, "right");
const toIdx = slide.shapes.getConnectionSiteIndex(targetShape, "left");

slide.shapes.connect(sourceShape.id, targetShape.id, {
  fromIdx,
  toIdx,
  kind: "straight",
  line: { style: "solid", fill: "slate-700", width: 2 },
});
```

Connection site indexes are PowerPoint preset-geometry connection points. A
connector's endpoint identity is the pair `(target shape, site index)`, not the
target alone. Keep the index when cloning, editing, or auditing a connector.

## Direct Connector Shape

```ts
const connector = slide.shapes.add({
  geometry: "connector",
  kind: "curved",
  from: sourceShape,
  fromIdx: 3,
  to: targetShape,
  toIdx: 1,
  line: { style: "solid", fill: "accent1", width: 2 },
  tail: { type: "triangle", width: "med", length: "med" },
});
```

Direct connector creation requires `from`, `to`, `fromIdx`, and `toIdx`. Prefer
`slide.shapes.connect(...)` unless exact connection site indexes are already
known.

## Direct Connector Inline Type

```ts
type ConnectorConfig = {
  geometry: "connector";
  from: Shape | string;
  fromIdx: number;
  to: Shape | string;
  toIdx: number;
  kind?:
    | "straight"
    | "elbow"
    | "elbow2"
    | "elbow3"
    | "elbow4"
    | "elbow5"
    | "curved";
  line?: LineConfig;
  head?: LineEndConfig;
  tail?: LineEndConfig;
  cap?: "flat" | "round" | "square";
  join?: "round" | "bevel" | "miter";
  accessibility?: { title?: string; description?: string; decorative?: boolean };
};
```

Connectors support line styling and line ends, but not `borderRadius` or
`shadow`.

## Edit Existing Connectors

```ts
const connector = presentation.resolve(connectorAnchorId);
const nextFromIdx = slide.shapes.getConnectionSiteIndex(
  nextSourceShape,
  "right",
);
const nextToIdx = slide.shapes.getConnectionSiteIndex(nextTargetShape, "left");

connector.setConnectorFrom(nextSourceShape, nextFromIdx);
connector.setConnectorTo(nextTargetShape.id, nextToIdx);
connector.line = { style: "solid", fill: "slate-800", width: 2 };
if (connector.accessibilityCapability.editable) {
  connector.setAccessibilityMetadata({ description: "Updated dependency flow." });
}
```

Use `presentation.inspect({ kind: "connector", search })` to find connector anchor
ids. Connector facades expose `connector`, `connectorLineStyle`,
`connectorHead`, `connectorTail`, `accessibility`, and
`accessibilityCapability` for readback. Non-visible metadata maps to
`p:nvCxnSpPr/p:cNvPr` and remains independent of the line name and geometry.
Its presence-aware `decorative` boolean uses the Office 2019+ extension;
`true` cannot coexist with title/description. Classification and corresponding
text clears/additions must be one transaction. Unknown or duplicate native
extension topology remains source-owned and fails closed on this mutation.

For a site-bound connector, do not assign `start` or `end` coordinates directly;
use `setConnectorFrom` or `setConnectorTo` so target and site change atomically.
The low-level `slide.connectors.add(...)` point/center API remains available for
legacy source-free workflows that intentionally have no exact site identity.

## Routing And Ordering

New connectors are sent behind shapes by default so boxes and labels remain
readable. Call `connector.bringToFront()` when the connector should sit above
other elements, or `connector.sendToBack()` to restore the default. Those
z-order methods are source-free operations. Imported connector z-order remains
source-bound and rejects instead of reordering an unmodeled SlidePart tree.
An imported connector with irregular `p:cNvPr` attributes or children keeps
those source bytes during unrelated supported endpoint/line edits, while a
metadata edit fails closed.

Connected routes update when a modeled endpoint shape moves. Render and export
paths recompute the connector bounds and route from the current endpoint
geometry. An unchanged imported connector may preserve a connection to an
unmodeled target, but moving or rewiring that target fails closed rather than
substituting its center.

## Native and imported boundary

OfficeKit reads and writes `straightConnector1`, `bentConnector3`, and
`curvedConnector3`; `elbow2` through `elbow5` normalize to the bounded elbow
model. It retains `a:stCxn/@id+@idx` and `a:endCxn/@id+@idx`,
solid/dashed/dotted/dash-dot/dash-dot-dot/no-line paint,
flat/round/square caps, round/bevel/miter joins, and bounded
triangle/stealth/diamond/oval/arrow ends with small/medium/large dimensions.
`dash`, `dot`, `dashDot`, and `longDashDotDot` normalize through the same line
profile used by ordinary shape outlines.

A recognized imported connector can change its bounded endpoints or line
profile while the surrounding element topology stays fixed. Source-bound
z-order is not editable. Missing connection IDs/indexes, duplicate connection
nodes, unsupported connector presets, custom adjustment graphs, theme/effect
outlines, or other unmodeled XML remain opaque/read-only and an attempted
semantic edit fails closed. No fallback rebuild substitutes a visually similar
line.

## Cookbook

```ts
// Arrow from one card to another.
slide.shapes.connect(sourceCard, targetCard, {
  kind: "elbow",
  fromSide: "right",
  toSide: "left",
  line: { style: "solid", fill: "slate-500", width: 2 },
  head: { type: "arrow", width: "med", length: "med" },
});
```

```ts
// Bidirectional relationship line.
slide.shapes.connect(leftShape, rightShape, {
  kind: "straight",
  fromSide: "right",
  toSide: "left",
  line: { style: "dashed", fill: "accent2", width: 2 },
  head: { type: "arrow", width: "sm", length: "sm" },
  tail: { type: "arrow", width: "sm", length: "sm" },
});
```

```ts
// Curved connector with explicit preset connection sites.
slide.shapes.add({
  geometry: "connector",
  kind: "curved",
  from: sourceShape,
  fromIdx: slide.shapes.getConnectionSiteIndex(sourceShape, "bottom"),
  to: targetShape,
  toIdx: slide.shapes.getConnectionSiteIndex(targetShape, "top"),
  line: { style: "solid", fill: "accent1", width: 2 },
  tail: { type: "triangle", width: "med", length: "med" },
});
```
