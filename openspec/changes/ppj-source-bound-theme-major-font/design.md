## Context

The authored compiler already stores `PresentationThemeArtifact.major_font_family` and lowers it to `a:majorFont`. Imported projections currently expose only a placeholder theme name and reject `fontScheme` edits because the imported ThemePart is source-owned. The existing source-preserving writer already has changed-part and opaque-content proofs.

## Goals / Non-Goals

**Goals:**

- Read one canonical major Latin typeface from a validated shared ThemePart.
- Reuse the existing wire field and bind a theme-level capability to `fontScheme.major`.
- Patch the typed Latin typeface attribute in place while preserving the remaining theme graph.

**Non-Goals:**

- Editing minor, East Asian, or complex-script typefaces in this increment.
- Creating missing font nodes, choosing among multiple ThemeParts, or changing theme relationships.
- Modeling theme fallback, inheritance, font embedding, or host font substitution.

## Decisions

1. **Require one shared ThemePart plus major and minor Latin values.** The minor value is retained in the projection because the PPJ font-scheme object requires it; requiring both avoids inventing a fallback for an incomplete source graph. A missing or ambiguous owner receives no font capability.
2. **Keep one theme native reference with a field-qualified capability.** The existing theme-level native reference can carry the new `setThemeFontScheme` operation alongside `setThemeName`; `fontScheme.major` is the only newly writable field.
3. **Use the existing artifact field.** Import fills `MajorFontFamily`, compilation changes that field, and the PPTX writer applies it after re-proving the canonical ThemePart. No protobuf field or protocol version changes.
4. **Write only the typed Latin attribute.** The writer updates `MajorFont.Latin.Typeface`, saves the ThemePart, records its path and replacement hash, and lets the existing opaque graph guard prove all untouched parts byte-for-byte.

## Risks / Trade-offs

- [Risk] Some producers omit a Latin font or attach different themes to masters. → Do not issue the capability and keep the graph source-owned.
- [Risk] A font name can be syntactically valid but unavailable on a host. → The field records the source XML value only; host font availability and rendering remain outside the claim.
- [Risk] The PPJ object exposes observed unowned slots beside one writable slot. → Compiler equality checks reject every change except `fontScheme.major` and the capability lists only that path.

## Migration Plan

No migration is required. Existing source-bound programs without the optional capability remain valid. Reverting the change restores the previous opaque `fontScheme` behavior.
