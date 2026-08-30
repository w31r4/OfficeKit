# Design

## Language

An arc command is written after a command that establishes a current point:

```json
{
  "op": "arcTo",
  "radiusX": 50,
  "radiusY": 30,
  "startAngle": 180,
  "sweepAngle": 180
}
```

Radii are positive finite values in the geometry view-box coordinate system.
Angles are degrees. `startAngle` is bounded to one signed turn and
`sweepAngle` is a non-zero signed value no greater than one turn; positive
sweep follows DrawingML's clockwise page-coordinate convention.

The current point supplies the point on the ellipse at `startAngle`, matching
the native DrawingML arc primitive. A preceding `moveTo` is the clearest
authoring form. Multiple paths and opposite sweep directions can form hollow
rings without flattening them to images.

## Native lowering and projection

NativeAOT converts view-box radii to the existing literal custom-path units and
degrees to DrawingML 1/60000-degree values, then writes `a:arcTo`. No new wire
field is required because protocol 2 already contains the typed arc command.

Projection accepts only literal arcs in the same bounded custom-geometry
profile. It normalizes a native start angle into one turn, preserves the signed
sweep and emits typed PPJ. Formula-backed arcs continue to make the geometry
opaque. Embedded PPJ recovery remains exact.

## Validation

Schema validation owns numeric ranges. Semantic validation rejects an arc
without an established current point before compilation. The native codec
independently revalidates positive radii, bounded sweep and command ordering.

