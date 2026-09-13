## Context

The source-bound theme projection already reads optional East Asian typefaces
into `PresentationThemeArtifact`, while the writer only exposes Latin major and
minor edits. This increment binds the existing major East Asian wire field to a
typed `a:majorFont/a:ea/@typeface` splice.

## Goals / Non-Goals

**Goals:**

- Bind `design.theme.fontScheme.majorEastAsia` to an existing source value.
- Reuse the theme native reference and existing artifact field.
- Change one font-scheme slot in one source-bound compile while preserving the
  rest of the package.

**Non-Goals:**

- Editing minor East Asian or complex-script faces, Latin slots in the same
  transaction, missing font nodes, multiple ThemeParts, fallback, inheritance,
  embedding, or host font substitution.

## Decisions

1. **Require the canonical owner and Latin scheme.** A single shared ThemePart
   with non-empty major/minor Latin values and an existing non-empty major East
   Asian value is required; otherwise no script-specific capability is issued.
2. **Use a dedicated capability identity.** The theme native reference keeps the
   existing major/minor Latin capabilities and adds
   `setThemeMajorFontEastAsia` for `fontScheme.majorEastAsia`, avoiding duplicate
   capability IDs and preserving the prior major/minor contracts.
3. **Keep edits field-level.** The compiler permits exactly one of major Latin,
   minor Latin, or major East Asian to change in a source-bound compile. Other
   font slots remain byte-preserving source-owned state.
4. **Write only the typed East Asian attribute.** The writer updates
   `MajorFont.EastAsianFont.Typeface`, records the ThemePart as changed, and
   lets the existing opaque graph guard prove untouched package bytes.

## Risks / Trade-offs

- Producers may omit the East Asian node or attach different themes to masters.
  → Keep the graph source-owned and issue no capability.
- A valid name may be unavailable on a host. → The field records source XML;
  host font selection and rendering remain outside the claim.
- The PPJ object exposes other observed font slots beside this owner. → Equality
  checks reject every unowned or combined change.

## Migration Plan

No migration is required. Existing source-bound programs without the optional
capability remain valid. Reverting the change restores the prior Latin-only
source-bound font behavior.
