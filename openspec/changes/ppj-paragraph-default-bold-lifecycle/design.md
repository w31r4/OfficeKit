## Context

See proposal.md. CollectTextLeafMutations masks paragraph alignment/tab stops only. Default-run properties already retain optional bold and support explicit modeled-style deletion.

## Goals / Non-Goals

Goals: source-bound defaultText.bold presence on ordinary text/shape owners with fixed rich-text topology.
Non-goals: changing other defaultText fields, slide-placeholder inheritance, table paragraph defaults or direct run styles.

## Decisions

- Add one field to setTextParagraphStyle capability projection and validation; require that exact field for bold edits.
- Mask only defaultText.bold during diff classification, pruning its empty wrappers. Retain rejection for all other changed styling/topology.
- Merge the requested bold presence into a clone of the imported default-run style. If no modeled fields remain, use the existing no_default_run_properties intent; other defaults and unknown native content remain source-owned.
- Update paragraph alignment/tab stops only when those fields changed, so a bold edit does not invoke unrelated deletion semantics.
- Reuse the shared PPJ lifecycle fixture with text and shape cases. Verify false/true/removal/restoration, bold-only wrappers, independent defaults/direct runs, untouched ZIP entries and missing authority.

- If the imported and requested default styles differ only in bold, patch a:defRPr/@b directly. The soft-edge sibling fixture exposed that rebuilding the whole default style unnecessarily invokes stricter effect writing; direct attribute editing preserves these siblings.

## Risks / Trade-offs

Default-run writing owns several existing fields; XML preservation comparisons must detect unintended sibling changes. Existing wire fields suffice, so no protocol regeneration is needed. Native host rendering remains separate.
