## Context

The imported presentation already has one canonical ThemePart and the six font typeface leaves are source-bound. OOXML also stores a human-readable name on the shared a:fontScheme element. It is a single attribute with a stable owner, so it can be exposed without opening the rest of the theme graph.

## Goals

- Represent the existing attribute as one optional PPJ string field.
- Keep source-bound edits limited to the existing a:fontScheme/@name attribute in the canonical ThemePart.
- Issue a capability for the exact field and reject edits that combine it with another theme font slot.
- Preserve source bytes and all other theme XML when the field is unchanged.

## Non-goals

- Creating a missing a:fontScheme or adding a name to an imported source that lacks one.
- Editing effect schemes, color schemes, font fallback, inheritance, or arbitrary theme XML.
- Claiming host font availability or PowerPoint rendering behavior.

## Decisions

1. Add font_scheme_name = 16 to PresentationThemeArtifact as an additive optional string. In the PPJ schema, fontScheme.name is optional and limited to 256 characters.
2. Source-free export uses the explicit font scheme name when supplied; otherwise the existing authored theme name remains the fallback. Imported source-free projection reads the written OOXML attribute.
3. Source-bound projection emits fontScheme.name only when the canonical a:fontScheme/@name exists and is non-empty, and adds setThemeFontSchemeName for fontScheme.name.
4. Source-bound compilation accepts one non-empty replacement value, requires the issued capability, and patches only that attribute. Removing it, combining it with another font slot, or using a forged capability fails closed.
5. A missing or ambiguous owner stays opaque: no field or capability is emitted and no compiler path creates topology.

## Data flow

a:fontScheme/@name -> PPTX import -> PresentationThemeArtifact.font_scheme_name -> PPJ design.theme.fontScheme.name plus field capability -> source-bound compiler -> the same ThemePart attribute.

## Risks and mitigations

The field is optional so older artifacts remain wire-compatible. The compiler checks the existing element and name before writing, and the regression compares changed parts and re-projects the result to catch accidental broad edits.
