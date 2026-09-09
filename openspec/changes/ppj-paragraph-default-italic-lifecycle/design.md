## Context

See proposal.md. Existing default-bold handling masks one field, checks exact authority and merges optional presence. Its native writer preserves unrelated font/effect children through a direct attribute patch.

## Goals / Non-Goals

Goals: default italic lifecycle on ordinary text/shape paragraphs, independently of bold.
Non-goals: other default fields, inherited placeholders, table paragraph defaults, host reflow.

## Decisions

- Extend field authority and diff masking with italic while checking bold and italic changes independently.
- Merge only changed optional flags; reuse existing default-style deletion when the last modeled field disappears.
- Compare default styles with both flags excluded, then patch only flags whose presence/value changed. Preserve the other flag's original XML spelling and all sibling children.
- Parameterize the existing lifecycle fixture across bold and italic, keeping soft-edge/unknown attribute/run/ZIP checks. Unsupported-field rejection moves to size now that italic is supported.

## Risks / Trade-offs

The known native default-style baseline failure stays documented separately. Updated codec is needed for new field authority; protobuf remains unchanged.
