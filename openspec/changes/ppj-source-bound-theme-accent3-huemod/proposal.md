## Why

F-15 already exposes bounded authored color transforms, but an imported accent3 theme color still has no independently editable hue modulation leaf. A strict source-bound field closes one small, auditable gap without pretending that the full theme transform graph is writable.

## What Changes

- Add `design.theme.accentTransforms.accent3.hueMod` as a bounded 0..1 fraction.
- Expose a `setThemeAccent3HueMod` capability only for one canonical shared ThemePart whose accent3 color is a direct RGB leaf with exactly one direct `a:hueMod/@val` child.
- Project the existing native value, edit only that token, and re-project the edited value.
- Reject missing or ambiguous owners, extra transforms, deletion, combined edits, out-of-range values, and capability tampering.
- Add the schema, registry, generated/reference documentation, backlog evidence, and one focused source-bound regression.

## Capabilities

### New Capabilities

- `presentation-theme-accent3-huemod`: A strict source-bound PPJ hue modulation leaf for the accent3 theme color.

### Modified Capabilities

None.

## Impact

The PPJ schema, capability registry, presentation projector/compiler/semantic validator, PPTX theme codec, presentation Skill reference, coverage/backlog docs, and NativeAOT codec tests are affected. No wire version change or new dependency is needed; unsupported theme topology remains source-owned.
