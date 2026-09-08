## Context

The completed formal text precedence slice distinguishes run, paragraph,
element, named-style, layout, master, theme, and default sources, but the
canonical PPJ `theme` currently contains colors only. The existing authored
compiler already lowers effective text scalars to native runs, while the
source-bound compiler deliberately rejects declaration graphs that do not have
an ownership proof.

## Goals / Non-Goals

**Goals:**

- Add one explicit `design.theme.textStyle` owner for the five formal text
  scalar fields.
- Make authored compiler and read-only review resolve the same theme value and
  fallback behavior.
- Preserve the existing native theme artifact and source-bound fail-closed
  policy.

**Non-Goals:**

- Do not model arbitrary `theme1.xml` font, effect, language, or color scheme
  inheritance.
- Do not write a new native theme node or issue a source-bound theme-edit
  capability.
- Do not change the Office wire version.

## Decisions

### PPJ representation

Add optional `theme.textStyle` using the existing direct `textStyle` schema.
Its values are direct fields rather than a second `defaultText` wrapper. This
keeps the new owner parallel to the authored direct `master.style` and
`layout.style` fields and avoids confusing it with the existing theme color
catalog.

### Shared precedence lookup

The existing formal text resolver will add the theme owner to its `theme`
branch. It continues to use the declared source order; a missing theme field
falls through to later sources, including a grammar default token. Existing
programs with no formal rule retain their historical inline/named-style path.

### Review agreement

For `text.*` rules, `reviewPpjArtifact` will read `design.theme.textStyle`
through the same direct-field helper used by the other text owners. It will
keep reporting one resolution entry per rich-text run and preserve the source
name `theme` in the evidence.

### Source-bound boundary

Semantic validation will reject `design.theme.textStyle` whenever the program
already has a source binding. This is explicit rather than relying only on a
generic raw JSON comparison, so callers receive a stable diagnostic. The
existing effective native theme output remains untouched.

## Risks / Trade-offs

- [Risk] A caller may read the PPJ theme fallback as a complete PowerPoint
  theme scheme. → [Mitigation] Keep the field documented as authored-only and
  reject it in source-bound programs; retain the full theme/effect boundary in
  the backlog.
- [Risk] Compiler and reviewer could disagree about omitted theme fields. →
  [Mitigation] Reuse the same direct-field/fallthrough shape and test both a
  theme winner and a default fallback.

## Migration Plan

Existing PPJ is unchanged. Authors may add `design.theme.textStyle` only when
they also declare a formal `text.*` precedence rule that includes `theme`.
Source-bound programs must omit the new field. Removing the field restores the
previous theme behavior.
