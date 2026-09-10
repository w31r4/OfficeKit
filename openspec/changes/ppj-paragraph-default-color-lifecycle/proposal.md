## Why

F-03 direct paragraph default color is projected and authored, but structured source-bound edits still reject it. Add an independent lifecycle while preserving source theme bindings, alpha presence and unrelated paint/effects.

## What Changes

- Support defaultText.color assignment, removal, color-only wrapper removal and restoration under exact field authority for ordinary text/shape paragraphs.
- Retain RGB/RGBA and color-token resolution, simple source-bound theme identity and explicit alpha presence including zero and one.
- Change only affected paragraphs, preserving other paragraphs' exact source alpha/theme bindings and direct run colors.
- Retain gradient ownership and reject color/gradient conflicts, unknown native paint and source luminance-transform replacement.
- Synchronize field guidance and add focused lifecycle/source tests.

## Capabilities

### New Capabilities
- `ppj-paragraph-default-color-lifecycle`: Direct paragraph solid-color presence and source fidelity.

### Modified Capabilities

None.

## Impact

PPJ capability schema/projector/validator, paragraph mutation, default-run writer, tests, Help and references. Existing RGB/scheme/opacity wire fields suffice. Gradient lifecycle, complex color transforms and host typography keep separate boundaries.
