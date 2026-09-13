## Context

The imported theme profile already proves one shared ThemePart and six strict direct RGB accent nodes. It projects all six observed values but currently permits only the accent1 writer path; the authored artifact model already carries the repeated six-value accent array.

## Goals / Non-Goals

**Goals:**

- Add one field-qualified capability for accentColors.accent2.
- Reuse the existing six-slot artifact and source ownership proof.
- Change only the existing accent2 val attribute and preserve no-op bytes.
- Add a focused source-bound regression and keep the OpenSpec contract synchronized.

**Non-Goals:**

- Editing accent1 through accent6 in one operation, or adding capabilities for the remaining slots in this increment.
- Creating missing nodes or accepting scheme colors, alpha, transforms, multiple ThemeParts, or inherited theme state.
- Changing the wire contract, protocol version, color-role model, effect scheme, font scheme, or host rendering behavior.

## Decisions

1. **Reuse the canonical six-slot projection.** The importer and projector already retain all six direct RGB observations. The new capability targets index one; accent1 keeps its existing capability, while accent3 through accent6 remain equality-checked source state.
2. **Add a separate capability name.** setThemeAccent2Color makes authority explicit and lets the semantic validator distinguish the new field from the existing accent1 capability. A generic capability would make field ownership ambiguous.
3. **Keep one-theme-field-per-compile enforcement.** The existing compiler rejects combined theme edits and unowned accent changes; accent2 participates in the same check instead of introducing a second mutation path.
4. **Patch the existing owner in place.** The native writer updates only a:accent2Color/a:srgbClr/@val after re-reading the six direct RGB nodes. It never creates a missing node or serializes the whole theme graph.
5. **Use the existing repeated wire field.** No protobuf or protocol change is needed. This keeps authored and imported paths compatible while the source-bound capability remains field-qualified.

## Risks / Trade-offs

- [Risk] Themes using scheme colors or transforms are common. → Keep them source-owned and do not issue the capability.
- [Risk] The projection shows observed values for four slots that are not writable in this increment. → Compiler equality checks reject changes outside accent2 and the native reference lists only authorized fields.
- [Risk] A new RGB value can render differently across hosts. → The regression proves XML/source preservation and second projection only; it does not claim host PowerPoint appearance.

## Migration Plan

No migration is required. Existing source-bound programs remain valid. Reverting this change removes only the accent2 capability and writer branch.
