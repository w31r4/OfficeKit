## 1. Character field lifecycle

- [x] 1.1 Align schema/native single-scalar validation, add exact character authority and patch the existing marker attribute; verify text/shape author/edit/restore, invalid scalars, required-field rejection and unknown-source/XML/ZIP preservation in focused tests.

## 2. Discoverability and shared behavior

- [x] 2.1 Update Help, schema descriptions, registry, generated PPJ manual/matrix, text guidance and F-03 backlog; verify direct character preview with explicit layout limits, related bullet/list tests, generated checks, portability/reference sync and strict OpenSpec validation.

Validation (2026-09-10): the focused cases first reproduced rejected source edits and accepted lone surrogates; after the fix 4/4 pass. They cover ordinary/supplementary/XML-escaped symbols, required-field and invalid-scalar rejection (including original JSON surrogate escapes), source no-op/edit/restore, unknown marker preservation and exact non-target XML/ZIP scope. Related paragraph, numbering, bullet/list/master and table-picture-bullet tests pass 26/26, zero skips. Direct character preview retains the symbol and partial layout evidence; missing font styling diagnoses unavailable. Generated checks, portability/reference sync and strict OpenSpec pass. No NativeAOT rebuild, full repository test or host glyph/layout acceptance was performed.
