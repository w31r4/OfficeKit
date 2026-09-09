## Why

F-03 has a direct Latin paragraph font lifecycle, while defaultText.fontFamilyEastAsia still rejects structured source-bound edits. Its deletion must also bypass authored Latin-to-East-Asian fallback.

## What Changes

- Add independent East Asian default font assignment, deletion, font-only wrapper removal and restoration for ordinary text/shape paragraphs.
- Honor explicit source-bound field presence without regenerating the authored fallback from Latin fonts.
- Preserve other script fonts, defaults, effects, direct runs and source XML; reject unmodeled East Asian font nodes.
- Extend the existing focused font lifecycle and rejection fixtures and public field guidance.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-east-asian-font-lifecycle`: Direct paragraph default East Asian font presence.

### Modified Capabilities

None.

## Impact

Capability schema/projection/validation, source paragraph mutation, default font writer, documentation and existing tests. Existing optional FontFamilyEastAsia wire field suffices; authored fallback and host font lookup remain unchanged.
