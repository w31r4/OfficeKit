## Why

F-07 still lacks direct chart character deletion-line state: a native `strike` attribute makes the shared chart style parser reject an otherwise supported owner. Ordinary PPJ text already expresses the three native strike states and boolean aliases.

## What Changes

- Add chartTextStyle.strike using the existing ordinary text vocabulary: true/false or noStrike/sngStrike/dblStrike.
- Preserve explicit noStrike separately from omission in native reading, writing, semantic comparison and source-bound edits.
- Propagate strike through existing chart style owners, rich trendline paragraph/run/end styles, grammar field precedence and vector text.
- Add focused lifecycle/invalid-value/wire preservation tests and update chart references and F-07 evidence.

## Capabilities

### New Capabilities
- `ppj-chart-text-strike`: Direct chart character strike state and its edit lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema, additive protobuf field/bindings, shared ChartML style codec, PPJ authored/source-bound builders/projector, vector text defaults, focused native/JS tests and chart documentation. No new dependency or host acceptance workflow.
