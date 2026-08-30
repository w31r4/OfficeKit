## Why

PPJ can author solid text color, highlight, typography and spacing, while its native DrawingML layer already has bounded gradient and outer-shadow codecs. The missing PPJ state makes a visible brand and title treatment impossible even though OfficeKit can represent it safely. This is a real language/compiler gap rather than an import limitation.

## What Changes

- Add bounded `gradient` and `shadow` properties to PPJ text styles used by direct runs and paragraph default text.
- Compile those properties into canonical DrawingML `a:gradFill` and `a:effectLst/a:outerShdw` state.
- Project the same canonical native state back into PPJ and preserve exact authored programs through embedded recovery.
- Reject solid-color/gradient conflicts and unsupported native fill or effect graphs instead of flattening them.
- Remove the text-effects authored compiler boundary and document the capability in the generated PPJ language reference and focused text guidance.

## Capabilities

### New Capabilities

- `ppj-text-effects`: Author, import and recover bounded gradient text and direct outer shadows through PPJ.

### Modified Capabilities

None.

## Impact

- Additive PPJ v1 schema and wire fields; Office wire version remains 2.
- C# text reader/writer, authored compiler and PPJ projector.
- Existing comprehensive PPJ contract, capability registry, generated `ppj.md`, focused text guidance and coverage evidence.
