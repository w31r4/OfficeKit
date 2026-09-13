## Context

Authored PPJ already lowers `accentTransforms` into direct DrawingML transform children, and the source-bound importer already treats transformed accent colors as opaque. A narrow imported profile can safely own one exact accent1 luminance modulation leaf without claiming the rest of the theme graph.

## Goals / Non-Goals

**Goals:**

- Expose `accentTransforms.accent1.lumMod` only for a canonical shared ThemePart with one direct RGB child and one direct `a:lumMod` child.
- Patch the existing luminance modulation `val` token in place and retain the complete source package outside that ThemePart.
- Reproject the edited fraction and enforce hash-bound, one-theme-field authority.

**Non-Goals:**

- Editing accent1 tint/shade, accent2 through accent6 transforms, `lumOff`, channel/discrete transforms, base accent colors, or any inherited/scheme/effect theme graph.
- Creating missing transform nodes, accepting alpha or extra topology, changing protobuf fields/protocol version, or claiming host color rendering.

## Decisions

1. **Reuse the existing transform wire model.** Store one `PresentationThemeColorTransform` with `Role = "accent1"` and `LuminanceModulationThousandth`; this keeps authored and imported representations aligned without a protocol change.
2. **Use a scalar PPJ capability.** Project only `accentTransforms.accent1.lumMod` and issue `setThemeAccent1LumMod`; the compiler accepts one fraction in `[0,1]` and converts it with the existing 100000 scale.
3. **Keep the source reader stricter than authored lowering.** The imported owner must be exactly `a:accent1Color/a:srgbClr` with only `val` plus one direct `a:lumMod/@val`; colors with other transforms, alpha, extra attributes, or descendants remain source-owned.
4. **Patch in place.** The writer rereads the canonical ThemePart, verifies the same direct topology, changes only `a:lumMod/@val`, saves that part, and records its hash. It never synthesizes a transform or rewrites the base color.
5. **Retain the existing mutation boundary.** A luminance modulation edit is one theme-field mutation; any simultaneous name, color, font, accent, role, tint, shade, `lumOff`, or other transform edit fails closed.

## Risks / Trade-offs

- Most real themes combine several transforms or use inherited colors, so they will intentionally receive no luminance modulation capability until a separate bounded profile proves those shapes.
- Values use the repository's existing `fraction * 100000` representation even though the wire field is named `Thousandth`; this matches authored lowering and preserves round-trip values.
- The focused regression proves XML/package preservation and reprojection, not PowerPoint host appearance or color-management fidelity.

## Migration Plan

No migration is required. Existing PPJ programs remain valid. Rolling back removes only the imported accent1 luminance modulation observation and writer branch; authored transforms and existing source-bound theme owners remain available.
