## Context

The native reflection message already stores optional `stPos` and `endPos`.
Direct rich-text runs now accept the bounded variable-start profile and issue
`textReflectionStartPosition`; the complementary variable-end profile can use
the same direct effect-list owner and text-leaf proof.

## Goals / Non-Goals

**Goals:**

- Read a direct run reflection with an explicit `endPos` and absent or zero
  `stPos` as a bounded semantic reflection.
- Issue one run/text-index-bound numeric leaf and splice only `@endPos` after
  proving the source token.
- Keep full-span run reflection scalars and the existing variable-start
  profile safe and source-preserving.

**Non-Goals:**

- Do not add protobuf fields, change authored reflection syntax, or change the
  wire version.
- Do not expose variable `stPos` together with variable `endPos`, transforms,
  WordArt, or host PowerPoint rendering as editable capabilities.

## Decisions

Use a dedicated `textReflectionEndPosition` leaf rather than treating the
position as opacity. Its binding stores the canonical `0..100000` token while
the PPJ value remains the existing `0..1` reflection position.

The direct run reader accepts variable positions only for a one-variable
endpoint profile: `stPos` must be absent or `0` for this leaf, while the
existing start leaf requires `endPos` absent or `100000`. A reflection with
both endpoints variable is rejected. For either variable endpoint, the five
legacy full-span scalar leaves are withheld; full-span runs retain those
scalars and any explicit endpoint leaves for lossless projection.

The source-bound writer reuses the text-leaf ordinal and direct
`rPr/effectLst/reflection` range, replacing only the `endPos` value. An
explicit full-span `endPos=100000` receives the dedicated leaf when the source
token is present. Malformed, noncanonical, stale, duplicate, or two-variable
reflections stay source-owned. The focused regression authors a direct run
reflection with `stPos=0` and `endPos=80000`, removes the embedded PPJ, edits
the leaf to `65000`, checks a SlidePart-only change and valid XML, and verifies
a second projection. A two-variable fixture proves the capability is not
widened.

## Risks / Trade-offs

- [Risk] Opting the run reader into variable positions could expose a
  two-variable graph. → [Mitigation] Accept only one variable endpoint and
  withhold legacy full-span leaves for that graph.
- [Risk] A semantically equal decimal could hide a noncanonical raw token. →
  [Mitigation] Validate the exact integer token and compare it during proof
  before patching.
