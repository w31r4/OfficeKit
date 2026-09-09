## Why

F-07 currently preserves chart reflection only when its alpha-ramp start/end positions are explicitly 0/100000. Chart styles need typed positions and native omission preservation to edit existing shorter or shifted reflection ramps without normalizing them to full span.

## What Changes

- Add chartTextStyle.reflection.startPosition/endPosition as optional 0..1 numbers with native thousandth-percent precision.
- Preserve both optional wire fields, explicit zero/defaults and absence through authored/source-bound chart styles, projection, named precedence and vector default text.
- Extend only the chart direct-reflection reader to variable positions. Ordinary text/shape/image imported full-span topology guards remain intact; their authored builders explicitly retain canonical positions.
- Update the minimal chart lifecycle and invalid-owner regressions and publish consistent OpenSpec, schema, wire and reference evidence.

## Capabilities

### New Capabilities
- `ppj-chart-reflection-positions`: independently editable chart reflection alpha-ramp positions.

### Modified Capabilities
None; prior chart reflection change remains unarchived, and this change extends its documented boundary.

## Impact

PresentationReflection wire fields 6/7, shared reflection value codec, chart builders/projector, ordinary canonical builders, PPJ schema, focused tests, capability references and F-07 backlog. No new ordinary PPJ syntax, partial-span native leaves, transforms or host acceptance claim.
