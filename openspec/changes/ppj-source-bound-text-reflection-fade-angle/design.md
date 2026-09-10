## Context

The reflection codec already parses and writes `fadeDir` when transforms are
enabled. Direct rich-text runs currently use the no-transform reader profile,
which deliberately keeps `fadeDir` source-owned. The existing five direct-run
reflection scalar leaves and the two endpoint leaves must remain unchanged.

## Goals / Non-Goals

**Goals:**

- Read one explicit canonical `fadeDir` on a full-span direct rich-text run.
- Issue a run/text-index-bound `textReflectionFadeAngleDegrees` leaf whose
  token is stored in native 60000ths of a degree.
- Replace only `reflection/@fadeDir` after the existing run/effect owner proof.

**Non-Goals:**

- Do not expose scale, skew, alignment, rotate-with-shape, or variable endpoint
  combinations in this change.
- Do not change authored syntax, the protobuf contract, wire version, or claim
  host PowerPoint rendering.

## Decisions

Use the existing reflection reader with transforms enabled only at the direct
run boundary, then apply a strict profile: `fadeDir` may be present, while
`sx`, `sy`, `kx`, `ky`, `algn`, and `rotWithShape` must be absent and the
reflection must have an absent/zero `stPos` and absent/100000 `endPos`. This
keeps the old endpoint profiles and full-span scalar leaves safe while allowing
one independently editable transform.

`textReflectionFadeAngleDegrees` uses the same canonical integer token as
`textDefaultReflectionFadeAngleDegrees`; PPJ exposes the token divided by
60000. The source-bound writer targets only `fadeDir` in the owning direct
`rPr/effectLst/reflection` range and compares the raw expected token before
splicing. Missing, malformed, noncanonical, stale, duplicate, or mixed
transform graphs stay source-owned or fail closed.

The authored compiler and PPJ schema accept the direct run's `fadeAngle` field.
The focused regression authors a direct run with `fadeAngle=45.5`, removes the
embedded PPJ, changes `fadeDir` from `2730000` to `5415000`, checks a
SlidePart-only change and valid XML, and verifies the second projection. A
reflection carrying an unsupported skew remains opaque.

## Risks / Trade-offs

- [Risk] Enabling transform parsing at the run boundary could widen the
  profile accidentally. → [Mitigation] Reject every transform except fadeDir
  and require full-span endpoint values before assigning the run reflection.
- [Risk] A typed value could hide a noncanonical raw token. → [Mitigation]
  Validate and compare the exact integer token during source-bound proof.
