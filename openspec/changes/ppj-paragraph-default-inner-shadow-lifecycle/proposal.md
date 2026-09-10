## Why

F-03 paragraph default inner shadow is authored and projected but has no independent source-edit lifecycle. Shared text-style syntax also forces blur, distance and angle into imported defaults, losing the distinction between omitted native geometry and explicit zero.

## What Changes

- Support defaultText.innerShadow assignment, removal, inner-shadow-only wrapper removal and restoration on ordinary text/shape paragraphs.
- Give paragraph default text its own schema view, reusing existing text property constraints; inner-shadow color remains required while geometry and opacity retain optional presence.
- Preserve RGB/RGBA/theme/token meaning, source grammar precedence, color transforms and native rounding.
- Patch only the direct innerShdw node, preserving mixed effect siblings, direct runs, neighboring paragraphs and non-target source parts.
- Keep duplicate/DAG/unmodeled source graphs fail-closed and synchronize discovery, guidance and focused regressions.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-inner-shadow-lifecycle`: Independent default inner-shadow state and optional geometry with source preservation.

### Modified Capabilities

None.

## Impact

PPJ schema, authored paragraph construction, default style projection, source-bound authority/lowering and targeted native writing; Help, capability metadata, references and tests. Reuse existing wire fields. Other inner-shadow owners retain their current schema and construction behavior.
