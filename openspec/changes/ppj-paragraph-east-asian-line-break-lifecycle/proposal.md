## Why

F-03 still lacks the direct setting controlling East Asian paragraph line-break
rules. PPJ cannot currently express its enabled, disabled or absent state.

## What Changes

- Add boolean `text.paragraphs[].style.eastAsianLineBreak` for East Asian line-start/end rules, independent of Latin word breaking and hanging punctuation.
- Preserve true, false and absence through authoring and snapshot-free native projection.
- Support ordinary text/shape source-bound add/set/remove/restore under exact field authority, retaining unknown native values and non-target content.
- Synchronize Help, registry, generated references and focused text guidance; verify minimal native experiments and explicit preview diagnostics.

## Capabilities

### New Capabilities

- `ppj-paragraph-east-asian-line-break-lifecycle`: direct East Asian line-break setting and source-preserving paragraph edits.

### Modified Capabilities

None.

## Impact

PPJ schema, paragraph wire binding, shared native paragraph codec, authored and
source-bound compilation/projection, documentation and tests. Inherited values,
kinsoku evaluation, font measurement and host layout remain separate evidence.
