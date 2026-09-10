## Why

F-03 paragraph default glow is authored/projected but lacks an independent source-edit lifecycle. Mixed effect lists also lose modeled defaults during projection, making glow deletion expose previously hidden siblings and fail writeback checks.

## What Changes

- Support defaultText.glow assignment, deletion, glow-only wrapper removal and restoration on ordinary text/shape paragraphs.
- Preserve direct radius and opacity presence, RGB/RGBA/theme/token colors and source-bound grammar precedence.
- Read each bounded direct paragraph default effect independently so glow edits retain neighboring effect semantics and native XML.
- Patch only the target glow; retain unrelated siblings, direct runs and other paragraphs, and reject duplicate/DAG or unmodeled glow owners.
- Synchronize guidance and focused source-preservation/negative fixtures.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-glow-lifecycle`: Independent paragraph default glow presence with stable sibling effects.

### Modified Capabilities

None.

## Impact

PPJ field authority, authored default-style effect construction, source-bound paragraph lowering, default-run effect projection/writing, schema/Help/registry/references and focused tests. Reuse existing glow wire fields. Other default effect editing and host rendering remain separate.
