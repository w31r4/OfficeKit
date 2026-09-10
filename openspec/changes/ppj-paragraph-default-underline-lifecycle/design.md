## Context

See proposal.md. The authored compiler maps single/double to sng/dbl and accepts the existing 18 native underline tokens. The native reader withholds underline when uFillTx/uFill/uLnTx/uLn children occur. ClearModeled already retains such source content.

## Goals / Non-Goals

Complete independent direct underline presence for ordinary text/shape paragraphs using existing optional wire state. Underline color/stroke effect graphs and inherited placeholder defaults retain their own contracts; host rendering is not established by XML tests.

## Decisions

Extend the current field authority, masking and cloned-default mutation path. Reuse alias lowering and canonical enum validation. Patch only a changed underline attribute in the scalar writer rather than rebuilding unrelated font/effect children.

Share the existing underline-effects predicate with the writer. Reject assignment when an unknown source token or any underline effect child exists, even without a u attribute. Unrelated scalar edits and removal must keep that original content; no-op keeps original bytes.

## Risks / Trade-offs

Aliases may differ from projection → assert native sng/dbl separately from projected single/double.
Explicit none may collapse into omission → check a cancellation/deletion chain and wrapper removal.
Effect-only source defaults may be overwritten → exercise effect children with and without u, alongside an unknown token fixture.
Legacy whole-default-style replacement has a documented baseline failure → retain that exclusion; this increment targets independent underline only.
