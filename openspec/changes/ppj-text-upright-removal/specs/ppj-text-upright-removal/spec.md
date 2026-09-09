## Purpose

Preserve the distinction between explicit true, explicit false and absent upright text-body state throughout PPJ authored and source-edit lifecycles.

## ADDED Requirements

### Requirement: Upright source state can be removed

Supported PPJ text, shape, owner-local placeholder and structured table-cell text styles SHALL retain authored true/false/absence. Removing a previously projected upright property SHALL remove the direct native upright attribute. Removing its otherwise empty style owner SHALL have the same effect. Explicit false SHALL remain a native value rather than removal. Existing source authority checks SHALL apply.

#### Scenario: Remove and restore upright
- **WHEN** a source request removes upright and a later fresh-source request adds true or false
- **THEN** export and fresh projection SHALL preserve absence and the newly requested value respectively

### Requirement: Upright edits preserve surrounding content

No-op SHALL preserve source bytes. Upright changes SHALL preserve unrelated body properties, text topology and non-target package members. Unsupported owners and unauthorized edits MUST reject; removing a style owner with additional properties MUST retain its prior rejection boundary.

#### Scenario: Owner-local edit
- **WHEN** an authorized source edit changes only upright
- **THEN** only the target owner's native upright state SHALL change and preview SHALL continue reporting unimplemented layout semantics
