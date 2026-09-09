## Context

See proposal.md. The codec currently recognizes straightConnector1, bentConnector3 and curvedConnector3, with at most a literal midpoint adj1. It stores endpoints but no adjustment presence. Curved authored output currently inserts a default literal.

## Goals / Non-Goals

Goals: preserve one native signed 32-bit adj1 value, including zero and absence, across authored/source PPJ operations. Non-goals: additional connector families, computed guide graphs, obstacle avoidance, arbitrary rotated bend axes or host appearance equivalence.

## Decisions

- Use bendAdjustment in native adjustment units, not points or a clamped ratio. Optional int32 wire presence distinguishes omitted from explicit zero; the default remains native geometry behavior.
- Recognize a single canonical `adj1` / `val N` without extensions. Reject formula/duplicate/unknown guide graphs. For nonstraight paths, rotations not equivalent to 0/180 degrees remain opaque because endpoint normalization cannot preserve their bend axis.
- Extend setConnectorType with bendAdjustment authority. Editing/removing the scalar replaces only recognized geometry; straight plus an explicit adjustment rejects. Changing type to straight clears an unchanged projected bend; an explicitly supplied conflicting bend rejects.
- Elbow/curved authoring without the field leaves avLst empty. Source no-op remains byte-identical; changed geometry retains endpoint/arrow/binding state and non-target ZIP members.
- Internal preview must diagnose non-default unpainted bends rather than display a midpoint route as correct. Existing production connector partial diagnostics remain in force.

## Risks / Trade-offs

- Zero confused with default -> optional wire presence and a zero regression.
- Rotation loses the independent bend axis -> retain unsupported sources as opaque.
- Imported literal graph partially reconstructed -> strict whole-owner recognition, reject computed/extended guides.

## Migration Plan

Additive optional wire tag and schema property, regenerate JS bindings, run proto check and focused connector/preview checks. Refresh PPJ source projections for new authority. No NativeAOT release or host gate is implied.
