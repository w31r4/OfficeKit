## Context

See `proposal.md` for the motivation. The existing custom-geometry reader
already recognizes ordered `a:cxnLst/a:cxn/a:pos` entries and retains literal
or reference coordinates. Generic PPJ native leaves are source-bound scalar
operations whose edit plan is proved against the original XML before a token
splice.

## Goals / Non-Goals

**Goals:**

- Give a direct literal x coordinate a stable native index and explicit EMU
  value.
- Reuse the existing native-leaf validation and SlidePart-only scalar splice
  contract.
- Prove that a changed x value survives a second projection while adjacent
  connection-site and custom-geometry data remains intact.

**Non-Goals:**

- Editing y coordinates, angles, formulas, handles, paths, or list topology.
- Evaluating a guide/reference or recalculating descendants after the edit.
- Adding wire fields or changing the Office protocol version.

## Decisions

1. **Use one leaf per ordered site index.** The index is already the native
   identity used for connection-site targeting, so no name inference or
   coordinate matching is needed. A pair-level position object was rejected
   because it would broaden the edit surface beyond the single x token.

2. **Require a canonical non-negative integer token.** The native model stores
   a resolved coordinate, but source-bound writing must distinguish a literal
   token from a reference and must preserve textual preconditions. Decimal,
   whitespace-variant, missing, negative, and out-of-frame values therefore do
   not receive the leaf. Formula evaluation was rejected because it would make
   a local x edit responsible for an opaque guide graph.

3. **Locate the position as a direct child and splice `@x`.** The writer will
   prove one direct custom geometry, one direct connection-site list, one
   direct position child per site, and only the allowed `ang`, `x`, and `y`
   attributes. It then changes the selected position's x attribute value and
   leaves the rest of the XML byte sequence and package relationships alone.

4. **Use the existing generic native-leaf contract.** The new kind is a
   numeric native leaf in the PPJ schema and registry. This avoids a protobuf
   descriptor change because the underlying connection-site coordinate is
   already present in the presentation model.

## Risks / Trade-offs

- [Risk] The model's resolved coordinate can hide whether the source used a
  literal or a reference. → Require the raw source token to be canonical and
  re-prove it in both the edit-plan proof and readback path.
- [Risk] A source geometry may contain legal but unmodeled extension content.
  → Reject non-direct or extra structure for this leaf and leave the original
  source opaque.
- [Risk] A coordinate change can alter rendered connection routing in a host.
  → Keep the capability scalar and source-bound; do not promise automatic
  connector rerouting or full host rendering equivalence.
