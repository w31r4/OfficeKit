## Why

F-03 paragraph default highlight is authored but lacks independent structured source-bound editing. Direct native theme highlights are also missing from PPJ projection; expose them without converting untouched paragraph styles to RGB.

## What Changes

- Add exact defaultText.highlight authority and independent assignment, deletion, highlight-only wrapper removal and restoration for ordinary text/shape paragraphs.
- Retain opaque RGB, color tokens and tint/shade through existing color resolution; project simple native theme highlights as token objects.
- Preserve standard source-bound theme tokens as native scheme colors when untransformed and not shadowed by grammar tokens.
- Change only paragraphs whose requested highlight differs; preserve other defaults, direct runs and source-owned highlight graphs.
- Add focused lifecycle/source fixtures and synchronize field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-highlight-lifecycle`: Direct RGB/theme paragraph default highlight lifecycle.

### Modified Capabilities

None.

## Impact

Schema authority, projector/compiler/default-run writer, source topology checks, tests, Help and references. Existing RGB/scheme optional wire state suffices. Nonopaque output, source color transforms/unknown topology and host typography retain their separate boundaries.
