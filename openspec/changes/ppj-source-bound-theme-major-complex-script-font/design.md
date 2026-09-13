## Context

The source-bound theme projection already reads optional complex-script
typefaces into `PresentationThemeArtifact`, while the writer exposes the Latin
and East Asian slots. This increment binds the existing major complex-script
wire field to a typed `a:majorFont/a:cs/@typeface` splice.

## Goals / Non-Goals

**Goals:**

- Bind `design.theme.fontScheme.majorComplexScript` to an existing source value.
- Reuse the theme native reference and existing artifact field.
- Change one font-scheme slot in one source-bound compile while preserving the
  rest of the package.

**Non-Goals:**

- Editing minor complex-script or other font slots in the same transaction,
  missing font nodes, multiple ThemeParts, fallback, inheritance, embedding,
  or host font substitution.

## Decisions

1. **Require the canonical owner and Latin scheme.** A single shared ThemePart
   with non-empty major/minor Latin values and an existing non-empty major
   complex-script value is required; otherwise no script-specific capability is
   issued.
2. **Use a dedicated capability identity.** The theme native reference keeps
   the existing Latin and East Asian capabilities and adds
   `setThemeMajorFontComplexScript` for `fontScheme.majorComplexScript`,
   preserving prior field contracts.
3. **Keep edits field-level.** The compiler permits exactly one of the existing
   font-scheme owners to change in a source-bound compile. Other font slots
   remain byte-preserving source-owned state.
4. **Write only the typed complex-script attribute.** The writer updates
   `MajorFont.ComplexScriptFont.Typeface`, records the ThemePart as changed, and
   lets the existing opaque graph guard prove untouched package bytes.

## Risks / Trade-offs

- Producers may omit the complex-script node or attach different themes to
  masters. → Keep the graph source-owned and issue no capability.
- A valid name may be unavailable on a host. → The field records source XML;
  host font selection and rendering remain outside the claim.
- The PPJ object exposes other observed font slots beside this owner. → Equality
  checks reject every unowned or combined change.

## Migration Plan

No migration is required. Existing source-bound programs without the optional
capability remain valid. Reverting the change restores the prior source-bound
font behavior.
