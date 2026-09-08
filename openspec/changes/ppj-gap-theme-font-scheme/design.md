## Context

`PresentationThemeArtifact` already carries the two authored family values,
and the PPTX writer already emits Latin, East Asian, and complex-script
typefaces for both native theme roles. The missing piece is an explicit PPJ
owner: `design.fonts[0]` and `[1]` are currently interpreted by position.

## Goals / Non-Goals

**Goals:**

- Make the major/minor theme roles explicit in the PPJ language.
- Preserve legacy programs that only declare `design.fonts`.
- Prove the declared roles in actual `a:fontScheme` XML.
- Keep source-bound theme graphs immutable and fail closed.

**Non-Goals:**

- Do not expose arbitrary imported `theme1.xml` children or theme-part edits.
- Do not add per-language fallback chains, font embedding, substitution, or
  PowerPoint host font evaluation.
- Do not change the Office wire version.

## Decisions

### Use a small explicit role map

`design.theme.fontScheme` contains required string fields `major` and `minor`.
Each value is a family name, not a font-catalog ID. This matches the native
theme role and avoids making array position part of the public contract.

### Keep one native bounded lowering

The existing `PresentationThemeArtifact` fields remain the wire owner. The
explicit role map takes precedence over the legacy font array; the writer
continues to place each role's family into its Latin, East Asian, and
complex-script slots. That is an intentional bounded profile, not a claim
about complete multilingual fallback.

### Source-bound remains read-only

Imported PPJ does not expose the source theme graph as an editable semantic
object. A source-bound program carrying `design.theme.fontScheme` is rejected
by semantic validation before any package mutation.

## Risks / Trade-offs

- Existing authors may have relied on the positional font array. The explicit
  field is optional and only overrides it when present.
- Consumers may mistake the three equal native script slots for complete font
  fallback. The schema/reference and test boundary document that host/script
  fallback remains outside the profile.

## Migration Plan

No migration is required. Existing programs continue to use the positional
`design.fonts` fallback. New authored programs may declare the explicit role
map when the theme scheme itself is part of the evidence.
