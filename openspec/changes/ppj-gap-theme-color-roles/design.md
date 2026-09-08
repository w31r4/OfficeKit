# Design

## Boundary

`design.theme.colorRoles` is an authored-only, fixed six-role map. Each
property is an opaque `#RRGGBB` value and maps directly to one of the native
DrawingML `a:clrScheme` roles outside `accent1` through `accent6`:
`dark1`, `light1`, `dark2`, `light2`, `hyperlink`, and `followedHyperlink`.

The map is deliberately narrower than a complete OOXML theme. It does not
model alpha, luminance/saturation transforms, effect styles, inheritance,
script-specific font fallback, or arbitrary theme XML.

## Lowering and recovery

The authored PPJ catalog lowers the six values into optional fields on
`PresentationThemeArtifact`. The existing source-free PPTX writer owns the
corresponding six `a:*` color elements and keeps its current defaults when a
role is omitted. The focused test inspects the six native XML values and
recovers the declaration through the embedded PPJ snapshot.

## Source-bound behavior

The imported theme graph remains source-owned. A source-bound PPJ that adds
`design.theme.colorRoles` is rejected with
`ppj.sourceBound.themeColorRoles` before any package part is written.

## Non-goals

- Reconstructing or editing arbitrary imported `theme1.xml` children.
- Modeling color transforms, transparency, effect schemes, or theme cascade.
- Changing the Office wire version.
