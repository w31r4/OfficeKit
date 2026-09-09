## Why

F-07 chart text still rejects native character properties carrying `lang`, including ordinary imported trendline rich text. PPJ already has a bounded language-tag contract for presentation text; chart text needs the same field and lifecycle.

## What Changes

- Add optional `chartTextStyle.language` using the existing languageTagOrToken contract and retain explicit native language presence.
- Propagate it through chart title, legend, axis, data-label and trendline text styles, grammar precedence, rich-text paragraph/run/end styles and vector title defaults.
- Reuse the pure language-tag validator in the shared codec assembly; retain the global-font-family-only profile boundary.
- Verify the narrow authored/source-bound/native/reprojection lifecycle and update chart documentation.

## Capabilities

### New Capabilities
- `ppj-chart-text-language`: Native chart character language state.

### Modified Capabilities

None.

## Impact

Schema, additive chart text-style protobuf field/bindings, shared native chart style codec, PPJ compiler/projector/precedence, two codec project include lists, focused tests and presentation references. No new dependency, dictionary download or host spellchecking requirement.
