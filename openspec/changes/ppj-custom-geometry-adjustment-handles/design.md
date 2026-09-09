## Context

See proposal.md. Native XY/polar handles support optional paired literal/reference ranges and required positions. The codec validates declared adjustment ownership, current value inside bounds and position inside shape. Source topology fixes order, kind and controlled names.

## Goals / Non-Goals

Expose both native handle records for custom shapes with literal paths. Keep full host drag behavior, referenced path coordinates and mask/clip owners separate.

## Decisions

Use a kind-discriminated `geometry.adjustmentHandles` list (max 1024). XY carries xAdjustment/minX/maxX and yAdjustment/minY/maxY; polar carries radialAdjustment/minRadius/maxRadius and angleAdjustment/minAngle/maxAngle. Both require position {x,y}. Coordinates/radii use shape-local points, angles degrees, independently of path viewBox; strings retain native references. Optional names are omitted when uncontrolled, ranges retain absence, empty lists project as omission. Reuse native graph validation and topology checks for source edits under a new setGeometry field. Do not flatten handles into numeric leaves.

## Risks / Trade-offs

Presence loss -> test absent ranges and zero. Unit confusion -> verify native EMU/angle values. Identity drift -> source topology rejects name/kind/order/list changes. Preview overclaim -> explicit handle diagnostic.

## Migration Plan

Additive PPJ schema and capability field; no wire change. Fresh projection gains editable literal-path shapes that have valid handles.
