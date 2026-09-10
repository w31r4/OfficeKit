## Why

F-03 paragraph default strike is authored and projected but lacks independent structured source-bound editing. Complete cancellation, removal and restoration without disturbing direct run styles or original text.

## What Changes

- Enable defaultText.strike assignment, explicit cancellation, deletion, strike-only wrapper removal and restoration for ordinary text/shape paragraphs.
- Preserve true/false aliases for sngStrike/noStrike and the noStrike/sngStrike/dblStrike enum; fresh projection uses canonical strings.
- Preserve other default/run styles and unmodeled source strike during no-op/unrelated edits; reject replacing unmodeled strike.
- Add focused regressions and synchronize field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-strike-lifecycle`: Independent direct paragraph default strike lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema/authority, paragraph diff/mutation, default-run writer, tests and references. Existing optional Strike suffices; no wire change. Host glyph rendering and inherited placeholder defaults retain separate boundaries.
