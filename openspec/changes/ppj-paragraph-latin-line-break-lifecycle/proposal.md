## Why

F-03 still lacks the direct setting that allows Latin text to wrap inside a word.
PPJ currently cannot express enabled, disabled or absent state for this setting.

## What Changes

- Add boolean `text.paragraphs[].style.latinLineBreak`, independent of text-box wrap and explicit line-break inlines.
- Preserve true, false and absence through authoring and snapshot-free native projection.
- Support ordinary text/shape source-bound add/set/remove/restore under exact field authority, preserving unknown native values and non-target content.
- Synchronize Help, registry, generated references and focused text guidance; verify minimal native experiments and explicit preview diagnostics.

## Capabilities

### New Capabilities

- `ppj-paragraph-latin-line-break-lifecycle`: direct Latin word-break setting and source-preserving paragraph edits.

### Modified Capabilities

None.

## Impact

PPJ schema, paragraph wire binding, shared native paragraph codec, authored and
source-bound compilation/projection, documentation and tests. Inherited values,
actual word breaking and host text layout remain separate evidence.
