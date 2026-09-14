## Why

F-15 has bounded source-bound ownership for accent1 and accent2 transforms, but an imported accent4 theme color still has no independently editable shade leaf. A strict source-bound field closes one small, auditable theme gap without turning the full theme transform graph into a writable object.

## What Changes

- Add `design.theme.accentTransforms.accent4.shade` as a bounded 0..1 fraction.
- Expose a `setThemeAccent4Shade` capability only for one canonical shared ThemePart whose accent4 color is a direct RGB leaf with exactly one direct `a:shade/@val` child.
- Project the existing native value, edit only that token, and re-project the edited value.
- Reject missing or ambiguous owners, extra transforms, deletion, combined edits, out-of-range values, and capability tampering.
- Add schema, registry, generated/reference documentation, backlog evidence, and one focused source-bound regression.

## Capabilities

### New Capabilities

- `presentation-theme-accent4-shade`: A strict source-bound PPJ shade leaf for the accent4 theme color.

### Modified Capabilities

None.

## Impact

The PPJ schema, capability registry, presentation projector/compiler/semantic validator, PPTX theme codec, presentation Skill reference, coverage/backlog docs, and NativeAOT codec tests are affected. No wire version change or new dependency is needed; unsupported theme topology remains source-owned.
