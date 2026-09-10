## Context

See proposal.md. PresentationTextParagraph currently has fields through 33. The shared paragraph codec models alignment and level using raw-token classification and per-property source edits; PPJ lacks a direction property. Source text mutation has explicit masks, per-field authority and per-paragraph patching. Ordinary rich-text projection requires supported inline topology.

## Goals / Non-Goals

**Goals:** Preserve native direction presence and make it independently writable in the existing typed paragraph profile.

**Non-Goals:** Unicode bidi resolution, glyph shaping, mirrored punctuation, inherited layout computation or widening unsupported text topology.

## Decisions

- Public direction uses left-to-right/right-to-left strings, consistent with columnDirection vocabulary but owned by the paragraph. Alignment remains independent.
- Add optional bool right_to_left = 34 to the existing paragraph wire. Presence distinguishes explicit false from absence without a new removal flag. Existing wire numbers and version stay unchanged; generate JavaScript with proto:generate and C# through its build.
- Model raw rtl tokens 0/1/false/true. Read and scrub only recognized tokens; unknown tokens stay residual. Requested presence may replace a recognized value or add a missing attribute, but cannot overwrite unknown state. Requested absence clears only a recognized direct attribute, matching alignment's source semantics.
- Lower style precedence through the existing paragraph path. Add exact direction authority and an independent source diff mask. Patch only changed paragraphs. Preserve equivalent native spelling on unrelated edits.
- Keep renderer support partial with an explicit unmapped native direction diagnostic, using the generated wire descriptor rather than a second field inventory.

The native mapping is documented by [Microsoft Open XML SDK](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textparagraphpropertiestype.righttoleft?view=openxml-3.0.1).

## Risks / Trade-offs

- False may collapse into absence → assert native attribute presence and fresh PPJ projection through add, set, remove and restore.
- New shared property could erase residual tokens or affect list/table/master defaults → test unknown token preservation and the existing shared paragraph regressions.
- Old codec binaries lack the new field → document that source and regenerated bindings are delivered here; release binaries require their normal build.
