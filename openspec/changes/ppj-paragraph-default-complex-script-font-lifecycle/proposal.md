## Why

F-03 paragraph defaults now have Latin and East Asian font lifecycles. The existing fontFamilyComplexScript field still lacks independent structured source-bound assignment and deletion.

## What Changes

- Authorize direct complex-script default font assignment, removal, font-only wrapper deletion and restoration for ordinary text/shape paragraphs.
- Preserve Latin/East Asian fonts, defaults, effects, direct runs and non-target source state.
- Reject extra metadata and hidden child content in source a:cs font nodes.
- Extend the shared font fixtures and field documentation.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-complex-script-font-lifecycle`: Direct paragraph complex-script default font presence.

### Modified Capabilities

None.

## Impact

Capability projection/schema/validation, source paragraph mutation, default font writer, references and tests. Reuse existing optional FontFamilyComplexScript; no wire or font-discovery changes.
