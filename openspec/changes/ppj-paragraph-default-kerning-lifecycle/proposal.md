## Why

F-03 paragraph default kerning can be authored and projected but lacks structured source-bound field authority. It needs independent presence and native precision semantics alongside the completed default scalar fields.

## What Changes

- Support defaultText.kerning assignment, explicit zero, deletion, kerning-only wrapper removal and restoration for ordinary text/shape paragraphs.
- Preserve other defaults, direct run kerning and unknown source content. Reuse finite 0..768pt validation and native hundredths with ties-to-even rounding.
- Add focused lifecycle/source preservation fixtures and synchronize public field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-kerning-lifecycle`: Independent direct paragraph kerning threshold lifecycle.

### Modified Capabilities

None.

## Impact

PPJ capability projection/schema/validation, paragraph mutation, default-run writer, tests and references. Existing optional FontKerningPoints suffices; no wire change. Host typography and inherited placeholder defaults remain separate.
