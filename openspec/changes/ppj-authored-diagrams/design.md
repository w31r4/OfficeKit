# Design

## Language

`smartArt` remains the public typed element for compatibility with PPJ v1. In
authored mode it represents a finite semantic diagram, not a promise to create
an Office SmartArt part:

```json
{
  "type": "smartArt",
  "id": "delivery-chain",
  "frame": { "x": 64, "y": 150, "width": 592, "height": 220 },
  "mode": "authored",
  "layout": "process",
  "shapeStyleRef": "diagram-node",
  "textStyleRef": "diagram-label",
  "connector": {
    "stroke": { "color": { "token": "ink-muted" }, "width": 1.5 },
    "endArrow": "triangle"
  },
  "nodes": [
    { "id": "observe", "text": "Observe" },
    { "id": "decide", "text": "Decide" },
    { "id": "act", "text": "Act" }
  ]
}
```

The element owns default `shapeStyleRef`, `textStyleRef`, and connector paint.
A node may override shape/text style refs. Picture layout nodes additionally
require an image asset. Explicit style refs are required for authored diagrams;
the compiler must not invent a palette or typography system.

## Layout profiles

- `list`: ordered vertical bands;
- `process`: ordered horizontal stages with forward connectors;
- `cycle`: ordered nodes around an ellipse with a closed connector sequence;
- `hierarchy`: parent graph arranged by depth, roots first;
- `relationship`: one center node with radial peers, or explicit parent edges;
- `matrix`: near-square row-major grid;
- `pyramid`: ordered centered bands that widen by level;
- `picture`: near-square image-and-label tiles.

Authored diagrams accept 1–64 nodes. Hierarchy requires a parent graph. Picture
requires one image asset per node. Other layouts use explicit parent edges when
present; process and cycle synthesize order edges only when no parent edges are
declared. Every frame, gap, shape, connector endpoint, and generated ID is a
deterministic function of the element frame, layout, and ordered nodes.

## Native lowering

The C# authored compiler emits one `PresentationGroup`. Connector children are
ordered before nodes so evidence labels remain visible. Nodes become ordinary
native shapes; picture tiles use an ordinary image plus a separate text shape.
The embedded PPJ remains the semantic authority, while every generated object
is still natively editable in PowerPoint.

Generated IDs use the element ID plus escaped node IDs. The compiler validates
the complete ID set before writing output. A second import of an OfficeKit deck
recovers the embedded PPJ exactly; an import without that snapshot sees the
ordinary group children rather than pretending they are a native SmartArt part.

## Source boundary

`mode: "source-bound"` is unchanged. Its nativeRef and node leaves continue to
target the original DiagramDataPart through the existing fail-closed edit plan.
Changing mode, layout, topology, IDs, or style on source-bound SmartArt remains
unsupported.
