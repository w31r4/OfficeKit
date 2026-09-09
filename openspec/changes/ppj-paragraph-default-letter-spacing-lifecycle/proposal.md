## Why

F-03 direct paragraph letterSpacing is authored and projected but lacks independent structured source-bound editing. Complete its presence and precision lifecycle alongside the other default scalar fields.

## What Changes

- Enable defaultText.letterSpacing assignment, explicit zero, deletion, spacing-only wrapper removal and restoration for ordinary text/shape paragraphs.
- Retain the finite -768..768pt range and native hundredths with ties-to-even rounding, including negative values.
- Preserve other defaults, direct run spacing and unknown native content; reject replacement of unmodeled native spc.
- Add focused source-preservation regressions and public guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-letter-spacing-lifecycle`: Independent paragraph default letter spacing lifecycle.

### Modified Capabilities

None.

## Impact

PPJ schema/authority, paragraph diff/mutation, default-run writer, tests and field documentation. Existing optional FontSpacingPoints suffices; no wire change. Host typography and inherited placeholders remain separate.
