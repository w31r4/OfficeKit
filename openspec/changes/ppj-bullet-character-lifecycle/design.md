## Context

See proposal.md. The schema currently allows 1–8 scalars; its validator already counts Unicode runes. `PptxBulletCodec` requires one enumerated rune but enumeration can replace lone surrogates, and it does not explicitly reject XML-incompatible characters. Source paragraph masks and native field capabilities omit `bullet.character`; native application rebuilds character markers even when retaining their kind.

## Goals / Non-Goals

**Goals:** Complete the required character value on existing character markers and preserve its native owner and all independent state.

**Non-Goals:** Marker-kind conversion, implicit symbol defaults, inherited bullet-style resolution or host font/glyph acceptance.

## Decisions

- Set schema scalar length to one and reject XML-incompatible BMP control/noncharacter values with a portable exclusion pattern. Validate exact rune consumption and XML character validity in the native codec so malformed UTF-16 cannot become a replacement scalar silently.
- Add an independent character mask and exact field capability. Lower only the changed paragraph's character on an existing character marker; retain the closed kind/topology guard.
- For character-to-character native application, retain the original element and assign Char only when its value changes. Rebuilding the marker would discard unknown attributes.
- Reuse the lifecycle helper for text/shape replacement and restoration, and add a small native invalid-scalar check plus parseable unknown-marker preservation. Exercise the existing direct-style SVG profile without changing the painter.

## Risks / Trade-offs

- UTF-16 length differs from Unicode scalar count → test a supplementary-plane symbol and reject a multi-scalar combining sequence.
- Invalid XML or lone surrogates can fail late or normalize → reject them in codec validation before serialization.
- Shared bullet code serves several owners → run focused bullet/list/master regressions and preserve different-kind replacement boundaries.
- Glyph availability depends on fonts → retain partial text-layout diagnostics and avoid a host-fidelity claim.
