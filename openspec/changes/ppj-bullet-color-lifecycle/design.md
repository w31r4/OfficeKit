## Context

See proposal.md. The native wire has RGB, scheme and follow-text choices with optional alpha. PPJ exposes a generic color; its RGB projection currently rounds alpha to hex bytes. The authored catalog resolves colors to RGB, while source paragraph edits already use independent masks and exact field authority.

## Goals / Non-Goals

**Goals:** Complete the direct choice without introducing a new wire field, preserving native theme/alpha state and source-local edits.

**Non-Goals:** Resolve full theme/list inheritance, change marker types, model imported arbitrary color-transform graphs or implement host glyph/layout acceptance.

## Decisions

- Add a bullet-specific color schema extending the existing color with RGB/optional-alpha object form. Keep hex output when its alpha exactly represents native state; otherwise emit the precise RGB object. This avoids silently rounding or changing unrelated color owners.
- Lower standard theme tokens directly when untransformed and not shadowed by grammar. Use a bullet-only grammar-preferred catalog resolution path for literal tokens/tint/shade; leave other catalog callers unchanged. Preserve explicit alpha even when it is one.
- Mask `color` and `colorFollowText` together, require authority for each actually changed property, and copy the requested wire color choice/alpha only to changed paragraphs.
- Extend shared color apply to remove a single modeled direct declaration for absence. Retain unchanged source color nodes using the existing semantic comparison. Unknown or duplicate choices remain untouched and reject replacement.
- Classify raw native color/alpha tokens before SDK enum/numeric reads, so malformed source values remain preservable. Recognize only direct RGB or scheme plus one valid alpha child.
- Use focused text/shape lifecycle fixtures, precise-alpha/theme/grammar examples, and malformed-source cases. Preserve font/size source spelling and non-target XML/ZIP; run related shared list/table/master tests and preview diagnostics.

## Risks / Trade-offs

- RGB alpha output can now be an object → document the typed form and prove author/project/re-author precision.
- Shared color deletion affects existing list/table/master writers → run their focused regressions.
- Theme/follow-text marker paint remains unresolved in preview → retain explicit unavailable diagnostics and host-layout limits.
