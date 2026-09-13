## Context

The authored theme transform model already stores alpha offset as `PresentationThemeColorTransform.AlphaOffsetThousandth` and writes direct DrawingML `a:alphaOff`. Source-bound theme editing recognizes only narrow, exact accent1 transform topologies. This change extends that boundary by one signed direct leaf and keeps the canonical shared ThemePart as the source of truth.

## Goals / Non-Goals

**Goals:**

- Recognize one direct accent1 RGB plus `a:alphaOff/@val` topology.
- Project a bounded signed `alphaOff` fraction and a hash-bound field capability.
- Compile one field mutation by token-splicing only the existing `a:alphaOff/@val`.
- Preserve all unrelated ThemePart and package content and fail closed on unsupported topology.

**Non-Goals:**

- Editing tint, shade, luminance, alpha modulation, channel/discrete/hue transforms, or multiple transforms in one request.
- Creating a missing transform, editing inherited/scheme/effect theme graphs, or claiming host color-management equivalence.
- Changing protobuf fields, protocol versions, or authored transform semantics.

## Decisions

1. **Use the direct alpha-offset leaf as the owner.** The reader requires exactly `a:accent1Color/a:srgbClr/a:alphaOff` with only the direct `val` attribute; this avoids inferring a writable owner from inheritance or a larger transform graph.
2. **Reuse the existing transform wire field.** Store the imported and edited integer in `AlphaOffsetThousandth`, mapping `-100000..100000` to the PPJ fraction `-1..1`; no protocol change is needed.
3. **Issue a scalar capability.** `setThemeAccent1AlphaOff` authorizes only `accentTransforms.accent1.alphaOff`, bound to the existing theme native reference hash.
4. **Patch in place.** The writer rereads and validates the canonical ThemePart, changes only the existing `a:alphaOff/@val`, saves that part, and records its hash. It never synthesizes a child or rewrites the base RGB.
5. **Retain the one-field mutation boundary.** Any simultaneous name, color, font, role, other transform, deletion, or capability change fails closed.

## Risks / Trade-offs

- **[Risk]** Real presentations may use alpha offsets with additional children, attributes, or inherited color owners. **Mitigation:** omit the field and retain the original topology as source-owned data.
- **[Risk]** PPJ fractions could round to a different integer. **Mitigation:** use the existing 100000 scale and midpoint-away-from-zero conversion, then reproject the stored integer.
- **[Risk]** A stale or modified native reference could authorize a different package. **Mitigation:** require the existing capability field and native reference hash before applying the patch.
