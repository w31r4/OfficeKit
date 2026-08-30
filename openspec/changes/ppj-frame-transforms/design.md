# Design

## One PPJ frame vocabulary

PPJ continues to use:

```json
{
  "frame": {
    "x": 72,
    "y": 96,
    "width": 320,
    "height": 180,
    "rotation": -6,
    "flipH": true,
    "flipV": false
  }
}
```

Arrays still define z-order. Rotation is expressed in degrees and compiled to
DrawingML 1/60000-degree units. Omitted values remove the canonical native
attribute for authored output; explicit false remains representable on import.

## Native ownership

`PresentationFrameTransform` is reused by Chart, Table, and Group. The codec
accepts only native `rot`, `flipH`, and `flipV` attributes within one revolution
in either direction. Group child coordinate space remains unchanged when the
outer frame rotates or reflects.

Shapes and images retain their existing wire transform messages, but
source-bound PPJ lowering now carries the same `frame` fields into them.

## Connector boundary

A connector's endpoints define its orientation. Independent frame rotation or
reflection would duplicate and conflict with those endpoints, so connector
`frame.rotation`, `flipH`, and `flipV` remain rejected. Opaque objects also stay
source-owned.

## Source-bound capability

For modeled Shape, Image, Chart, Table, and Group objects, `setFrame` declares
the seven bounded frame fields. Re-projection must restore them. Unsupported
native transform topology prevents the object from receiving that capability.
