## Why

F-03 still lacks the direct hanging-punctuation setting used for paragraph
typesetting. PPJ cannot currently express enabled, disabled or inherited state.

## What Changes

- Add boolean `text.paragraphs[].style.hangingPunctuation`, independent of hanging indent.
- Preserve true, false and absence through authoring and snapshot-free native projection.
- Support ordinary text/shape source-bound add/set/remove/restore under exact field authority, retaining unrecognized native state and non-target content.
- Synchronize Help, registry, generated manual/matrix and focused text guidance; verify minimal native experiments and explicit preview diagnostics.

## Capabilities

### New Capabilities

- `ppj-paragraph-hanging-punctuation-lifecycle`: direct paragraph hanging punctuation and source-preserving editing.

### Modified Capabilities

None.

## Impact

PPJ schema and paragraph wire binding, shared native paragraph codec, authored
and source-bound compilation/projection, documentation and tests. Font metrics,
line breaking and actual host punctuation placement remain separate evidence.
