# Design

## Boundary

`design.theme.accentColors` is an authored-only, fixed six-role map. Each
property is an opaque `#RRGGBB` value and maps directly to the same-numbered
DrawingML accent role. The field is deliberately narrower than the complete
OOXML color scheme: dark/light, hyperlink, followed-hyperlink, transforms,
and effect schemes are not inferred or edited.

When the map is absent, the existing first six `design.theme.colors` values
remain the compatibility fallback. This keeps existing PPJ programs stable
without making the old positional convention the only new authoring surface.

## Lowering and recovery

The authored PPJ catalog writes the six values into the existing
`PresentationThemeArtifact.AccentRgb` repeated field. The existing PPTX writer
already owns the six `a:accent*Color` elements, so no wire or package model
change is needed. The focused test inspects those six native XML values and
recovers the declaration through the embedded PPJ snapshot.

## Source-bound behavior

The native imported theme graph remains source-owned. A source-bound PPJ that
adds `design.theme.accentColors` is rejected with
`ppj.sourceBound.themeAccentColors` before any package part is written.

## Non-goals

- Reconstructing arbitrary imported `theme1.xml` color schemes.
- Modeling alpha, luminance/saturation transforms, hyperlinks, or effect
  style inheritance.
- Changing the existing protobuf wire contract.
