## Why

PPJ v1 deliberately types a chart title as `textContent`, so an Agent can use
multiple runs, typographic contrast, East Asian fonts, color, emphasis, and
paragraph structure inside the title. The authored compiler currently rejects
every non-string title and the PPTX projector flattens an imported rich title
to one string. That makes the public language contract false and removes a
high-value visual primitive from data-heavy presentations.

## What Changes

- Add an additive structured title body to the Presentation chart wire model.
- Compile a rich PPJ chart title into native `c:title/c:tx/c:rich` DrawingML.
- Recover the bounded native title body when importing a PPTX instead of
  flattening it.
- Preserve the existing string title and uniform `titleTextStyle` behavior.
- Permit source-bound rich-title edits only when the native rich text is inside
  the compiler-owned profile; retain the rest as opaque chart content.
- Synchronize generated PPJ guidance, chart guidance, capability ownership,
  and coverage evidence.

## Capabilities

### New Capabilities

- `ppj-rich-chart-title`: Authored and safely source-bound structured chart
  titles using PPJ paragraphs and runs.

### Modified Capabilities

None. The PPJ schema ID and Office wire protocol version remain unchanged.

## Impact

- The additive Presentation chart wire message, PPJ lowering/projecting, chart
  part reader/writer, generated documentation, capability registry, and one
  existing PPJ contract are affected.
- No raw OOXML, chart-layout DSL, arbitrary formula, or host automation surface
  is introduced.
