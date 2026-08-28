# Presentation template reference decks

## Why

Image-only style references can communicate a visual system but cannot expose
native layer order, editable objects, or reusable Office assets. A published
OfficeKit presentation template therefore needs one small, self-authored PPTX
calibration deck in addition to its complete style Skill and visual examples.

## What changes

- Add presentation template schema v4 with a required `assets/reference.pptx`.
- Make the Template Creator accept only an OfficeKit-authored, structurally
  valid reference deck for presentation packaging and bind its SHA-256.
- Return the authored reference path from local template search.
- Keep external decks and source images task-scoped evidence; never publish
  them or turn the reference deck into a fixed layout registry.
- Migrate templates in later slices, marking visual and functional restoration
  indices separately. A template is restored only when both are at least 95.

## Non-goals

This change does not copy Kimi or other third-party decks, add a second DSL,
clone fixed pages, or claim that a single reference deck covers every PPTX
feature.
