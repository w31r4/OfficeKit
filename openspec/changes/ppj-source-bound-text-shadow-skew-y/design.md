## Context

The committed direct-run shadow profile already parses `sx`, `sy`, and `kx` as optional transforms while restricting the effect list to a single direct outer shadow. DrawingML stores vertical skew in `a:outerShdw/@ky` as a signed integer in 1/60000-degree units. The existing source-bound editor already token-splices scalar shadow attributes and can preserve the owning SlidePart.

## Goals / Non-Goals

**Goals:**

- Add one PPJ field and one native leaf for direct-run outer-shadow vertical skew.
- Preserve canonical signed tokens, strict bounds, and single-transform safety.
- Prove authored round trip, source-bound token replacement, changed-part scope, and fail-closed topology.

**Non-Goals:**

- Paragraph `defaultText` shadows, chart text shadows, shape/image shadows, WordArt, or host rendering.
- Rebuilding an effect list or normalizing unsupported XML.

## Decisions

- Extend the shared shadow safety predicate with an explicit `allowSkewY` flag. The direct-run projection permits `ky` only as the one allowed transform and keeps `sx`, `sy`, and `kx` mutually exclusive through the existing single-transform proof.
- Store the PPJ value in degrees and the native leaf value as the canonical signed `ky` token. The editor maps the leaf to `ky` and performs a source-bound token splice rather than serializing the whole effect list.
- Keep the existing parser's strict geometry/color/unknown-node checks. This bounds the new capability to a directly provable owner and avoids broadening ordinary shape/image shadow support.

## Risks / Trade-offs

- [Risk] Native `ky` may be malformed, out of range, or combined with another transform. → Mitigation: do not issue the leaf or edit permission unless the strict owner proof passes.
- [Risk] Degree conversion can produce a boundary value after rounding. → Mitigation: validate the rounded signed integer and enforce the strict -90/90-degree range in both authored and source-bound paths.
- [Risk] A source-bound rewrite could disturb unrelated package bytes. → Mitigation: patch only the recorded attribute span and assert owning-SlidePart-only changes in the focused fixture.
