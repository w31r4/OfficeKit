## Context

See proposal.md. The wire already has repeated tab stops and a removal-intent flag. Native validation accepts nonnegative signed-32-bit EMUs and strictly increasing positions. PPJ currently advertises a smaller range, independently selects the two style properties, rejects ordinary removal and updates all paragraphs when any tab list changes. The native writer always replaces requested lists, while source classification ignores some unknown metadata and can parse malformed SDK values too early.

## Goals / Non-Goals

**Goals:** Complete the existing field lifecycle and preserve every untouched paragraph's native spelling and metadata.

**Non-Goals:** No new wire shape, inherited-tab solver or host font/layout engine. Empty arrays and noTabStops remain removal aliases rather than authored empty native lists; untouched imported empty lists remain source-preserved.

## Decisions

- Reuse the existing removal flag when a changed PPJ list becomes absent/empty. Keep native unset-as-preserve behavior for callers outside the PPJ mutation path.
- Select tabStops/noTabStops as one choice per style layer, so a direct clear suppresses a named list. Require exact authority for each changed property and patch only paragraphs whose choice changes.
- Parse a complete direct list using raw namespace-free pos/algn attributes before SDK conversions. Accept only known nodes/attributes, bounded count and increasing integer positions. Invalid lists stay outside modeled tab state while other paragraph fields remain editable; scrubbing removes only modeled lists.
- Compare modeled requested stops with native values before replacement. Equality preserves plus signs, padding, optional default alignment and unrelated edits without rewriting XML.
- Match PPJ positions to 0 through int.MaxValue/12700 points; validate order after nearest-even EMU rounding. Limit lists to 32 stops, matching the DrawingML CT_TextTabStopList schema and the Office2021 validation experiment.
- Use text/shape lifecycle fixtures with a decorated neighboring list, numeric/alignment negatives and source XML/ZIP comparisons. Reuse existing field/break/tab, list/table/master regressions; preview verifies the existing explicit limitation.

## Risks / Trade-offs

- Relaxing whole-paragraph rejection for unmodeled tab state could erase residual data → classify, omit and scrub the list atomically, and verify unknown XML through unrelated edits and refused replacements.
- Empty arrays are not retained as native list presence → document their existing clear intent; untouched imported empty lists remain byte-preserved.
- Shared tab helpers serve list defaults and tables → run their narrow existing tests; keep wire removal semantics unchanged.

The removal-intent flag must bypass semantic no-op dispatch for ordinary shapes and nested groups, including source clone refusal. It remains excluded from semantic hashes, and the native writer validates the actual list before removal.

Format limit source: [Open XML SDK DrawingML schema](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/schemas/schemas_openxmlformats_org_drawingml_2006_main.json), CT_TextTabStopList permits at most 32 tab children. The prior 256-entry implementation produced schema-invalid output.
