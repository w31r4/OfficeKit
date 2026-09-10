## Context

The existing direct rich-text reflection reader accepts the base geometry and the recently bounded endpoint and fade-angle profiles, but rejects all scale transforms. The protobuf model, native writer, and chart/default-text paths already carry a signed 1/100000 scale value. Source-bound edits use proof-bound raw XML token splices and must preserve unsupported topology.

## Goals / Non-Goals

**Goals:**

- Add one direct-run scaleX field with authored, projected, leaf, proof, and token-splice paths.
- Require full-span positions and no other reflection transform for the direct-run profile.
- Keep the smallest possible changed-part footprint and a second-projection regression.

**Non-Goals:**

- ScaleY, skewX, skewY, alignment, rotateWithShape, or combinations with fadeAngle.
- Paragraph default-text, shape, image, chart, host PowerPoint rendering, or arbitrary reflection graph edits.
- New protocol fields or lossy flattening.

## Decisions

Use textReflectionScaleX as the native leaf kind and run.style.reflection.scaleX as its PPJ location, with ratio values scaled by 100000. Extend direct-run safety checks so scaleX is accepted only for a full-span reflection with no fade or other transform; existing endpoint and fade profiles remain mutually exclusive with scaleX. The projector emits scaleX only when the native presence bit is set.

Reuse the existing reflection proof and raw XML patch path, adding only the sx attribute mapping and scalar validation. This keeps the source hash, text index, and owner-local token checks identical to the already-proven reflection fields. A separate generalized transform model would expose combinations that the source writer cannot prove, so it is deliberately rejected.

## Risks / Trade-offs

- [Unsupported topology] A valid PowerPoint reflection may remain opaque when it combines transforms; this is the intentional fail-closed boundary.
- [Precision] PPJ ratios round to native 1/100000; the authored and reprojected values are checked at that precision.
- [Host rendering] The regression proves package and projection behavior only; host PowerPoint display is not claimed.

## Migration Plan

No migration is required. Existing PPJ remains valid. Consumers that see no textReflectionScaleX leaf continue to treat the imported run as source-owned.
