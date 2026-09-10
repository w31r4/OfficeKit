## Why

F-03 paragraph layout currently exposes left indent and hanging indent, while the native right margin stays outside PPJ. Paragraph direction alone cannot express this independent right-side inset.

## What Changes

- Add `text.paragraphs[].style.rightIndent` in points, from 0 through 4032, rounded to the nearest EMU with ties to even. It always addresses the physical right paragraph margin, independently of direction, left indent, hanging indent and text-box insets.
- Extend the paragraph wire with a right-margin value/removal choice using new field numbers. Keep existing fields and wire version; generate bindings with the repository script.
- Support authored style precedence and ordinary source-bound text/shape add/set/remove/restore under exact rightIndent authority. Explicit zero retains the direct attribute; deleting the field or its rightIndent-only style wrapper removes it.
- Preserve equivalent native numeric spelling, neighboring paragraphs, runs and unrelated XML/ZIP. Invalid native right margins stay source-owned and reject replacement while other modeled edits remain available.
- Synchronize Help, schema, registry/manual/matrix, focused Agent guidance, backlog and explicit preview limitations.

## Capabilities

### New Capabilities

- `ppj-paragraph-right-indent-lifecycle`: independent right paragraph inset and its source-preserving lifecycle.

### Modified Capabilities

None. There is no corresponding main spec.

## Impact

PPJ schema and compilers/projector, paragraph wire/generated binding, shared paragraph layout/text codecs, presentation documentation and focused tests. No new dependency, automatic line reflow, inherited-margin solver or host typography acceptance.
