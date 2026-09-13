## Context

The source-bound theme projection already reads the canonical ThemePart's major
and minor Latin typefaces and writes the major slot. The artifact wire model
already contains `minor_font_family`, so this increment only needs the field
authority and typed `a:minorFont/a:latin/@typeface` splice.

## Goals / Non-Goals

**Goals:**

- Bind `design.theme.fontScheme.minor` to the existing minor Latin value.
- Reuse the theme native reference and existing wire field.
- Change one Latin slot in one source-bound compile while preserving the rest of
  the package.

**Non-Goals:**

- Editing major and minor together, East Asian or complex-script typefaces,
  missing font nodes, multiple ThemeParts, fallback, inheritance, embedding, or
  host font substitution.

## Decisions

1. **Require the same canonical owner as major.** A single shared ThemePart with
   non-empty major and minor Latin values is required; no capability is issued
   for an incomplete or ambiguous graph.
2. **Keep a separate field-qualified capability.** The theme native reference
   keeps the existing `setThemeFontScheme`/`fontScheme.major` entry and adds a
   `setThemeMinorFont`/`fontScheme.minor` entry. The compiler accepts exactly
   one of those two changed fields per compile, keeping each edit field-level
   with unique capability identities.
3. **Reuse the existing wire field.** Import already fills
   `PresentationThemeArtifact.minor_font_family`; source-bound compilation
   changes that value and the writer applies it after re-proving the canonical
   ThemePart.
4. **Write only the typed Latin attribute.** The writer updates
   `MinorFont.Latin.Typeface`, records the ThemePart as changed, and lets the
   existing opaque graph guard prove untouched package bytes.

## Risks / Trade-offs

- Producers may omit Latin faces or attach different themes to masters. → Keep
  the graph source-owned and issue no minor capability.
- A syntactically valid font name may be unavailable on a host. → The field
  records the source XML value only; host rendering remains outside the claim.
- The PPJ object exposes observed unowned slots beside the writable minor slot.
  → Equality checks reject every non-minor change and combined major/minor edits.

## Migration Plan

No migration is required. Existing source-bound programs without the optional
minor capability remain valid. Reverting the change restores the previous
major-only source-bound font behavior.
