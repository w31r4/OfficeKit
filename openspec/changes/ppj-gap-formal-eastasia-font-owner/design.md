## Context

The schema and native text model already represent `fontFamilyEastAsia`, and
authored compilation currently reads it through the legacy direct/inline
fallback chain. The formal resolver and reviewer already share the owner names
and lookup shape for five text scalars; only this sixth existing field is
missing from that profile.

## Goals / Non-Goals

**Goals:**

- Add one independent formal target for the East Asian typeface.
- Reuse the existing text owner lookup and theme text-style fallback.
- Keep compiler/reviewer source and value evidence aligned.
- Preserve legacy behavior when the target is not declared.

**Non-Goals:**

- Do not add language, complex-script, font embedding, font substitution, or
  host fallback semantics.
- Do not change native leaf capability names or wire messages.
- Do not make source-bound theme XML or inherited styles writable.

## Decisions

### Treat East Asian typeface as an independent scalar

The field has its own DrawingML owner and already appears in run styles and
native leaves. It therefore gets its own `text.fontFamilyEastAsia` rule rather
than being silently coupled to `text.fontFamily`. If it is absent after the
declared lookup, the existing compiler fallback from Latin font remains the
compatibility behavior.

### Reuse the existing direct theme text-style owner

`design.theme.textStyle.fontFamilyEastAsia` is already part of the bounded
authored theme fallback object. The formal theme branch reads that field just
as it reads the other text scalars; this does not imply arbitrary `theme1.xml`
font-scheme editing.

### Keep validation and review in lockstep

The target is added to the same semantic target set used to authorize
owner-bearing sources. The JavaScript reviewer already routes every `text.*`
field through its direct text-style lookup, so it only needs the existing
target to become test-visible; no separate font resolver is introduced.

## Risks / Trade-offs

- [Risk] Callers may interpret East Asian font selection as complete multilingual
  font fallback. → [Mitigation] Document and test it as one direct run scalar;
  language/script fallback remains outside the profile.
- [Risk] Compiler and review could diverge on missing values. → [Mitigation]
  Keep the same source order and verify both a theme hit and a default value.

## Migration Plan

No migration is required. Existing programs are unchanged. Authors that need
formal East Asian ownership add a `text.fontFamilyEastAsia` rule; removing the
rule restores the prior direct/inline behavior.
