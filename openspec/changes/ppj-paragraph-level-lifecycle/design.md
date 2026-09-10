## Context

The authored compiler already lowers level as an optional uint, and native paragraph validation bounds it to 0..8. Direct paragraph writing currently rewrites/removes `lvl` unconditionally; typed reads can throw on invalid source values. List/master level numbers also use this wire field as an identity, with direct attribute handling disabled.

## Goals / Non-Goals

**Goals:** Independent source-bound direct paragraph level edits and raw-attribute preservation.

**Non-Goals:** Changing list/master level identities, rebuilding inheritance, host reflow acceptance or a NativeAOT release.

## Decisions

- Add the exact source field authority, masked diff detection and changed-paragraph lowering following the existing paragraph scalar path.
- Reuse optional level presence as complete-paragraph deletion intent. A new no-level wire marker would duplicate the existing native paragraph contract.
- Parse native `lvl` as an invariant integer within 0..8. Reject assignment over unmodeled values; skip unchanged values and scrub only modeled direct attributes.
- Preserve `includeLevel`/`readLevel` boundaries and keep identity-only list/master levels out of the modeled-properties predicate.
- Reuse source/ZIP comparison helpers for text/shape fixtures, boundary values, defaults, lexical preservation, unknown values and exact authority. Run only affected paragraph/list regressions and documentation gates.

## Risks / Trade-offs

- Direct levels can select inherited list defaults → preserve source defaults and keep effective layout partial.
- Shared paragraph code serves list/master owners → retain their identity semantics and run their existing regressions.
- Invalid imported levels lack a typed representation → preserve them during unrelated edits and reject replacement.
