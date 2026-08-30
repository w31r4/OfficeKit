## Why

OfficeKit already owns bounded native chart-title font sizing and canonical line-chart smoothing/color-variation semantics, but the Presentation bridge and PPJ language do not expose them end to end. This leaves Agent-authored PPJ visibly less expressive than the underlying NativeAOT codec and creates another orphan capability that cannot survive build, import, and continued editing.

## What Changes

- Add chart-title text style to PPJ chart styles, initially owning explicit font size.
- Add bounded line-chart `smooth` and `varyColors` state to PPJ chart styles.
- Carry the existing native chart-title and line-options messages through the Presentation chart bridge.
- Compile, project, reimport, and document the new state without widening unsupported chart families or raw OOXML access.
- Keep source-bound edits fail closed unless the imported chart receives a dedicated capability for the changed style field.

## Capabilities

### New Capabilities

- `ppj-chart-typography-line-behavior`: Authored and recognized PPJ chart-title typography plus canonical line-chart smoothing and direct color-variation state.

### Modified Capabilities

None.

## Impact

- PPJ JSON Schema and generated language reference.
- NativeAOT PPJ authored compiler and PPTX projector.
- Presentation chart ↔ shared native chart-model bridge.
- Presentation Skill chart guidance and capability registry.
- One existing integrated PPJ build/reimport contract test.
