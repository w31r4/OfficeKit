## Why

OfficeKit can already create strong source-free decks and can describe imported
PPTX design evidence. The missing product path is to use that evidence to
generate new content in an existing visual system without flattening the
source package or pretending that every imported object is editable.

## What Changes

- Define template-conditioned generation as a distinct route from exact frame
  following and generic design presets.
- Use the bounded `Presentation.designProfile()` as descriptive evidence for
  canvas, typography, palette, density, archetypes, components, and opaque
  content.
- Use the read-only `Presentation.planTemplateGeneration()` primitive to turn
  page roles/content briefs into a source-bound multi-page frame map with
  clone-safe source slides, bounded text targets, reusable asset candidates,
  alternatives, and explicit blocked requests.
- Select codec-proven source slides, duplicate them through explicit
  export/reimport boundaries, and apply only inherited run/SVG-text edits.
- Reimport, verify, render when supported, and compare output issue categories
  against the source baseline before delivery.
- Freeze a three-sample real-asset benchmark and a source-free smoke fixture;
  record hashes and evidence, never vendor third-party PPTX files.

## Capabilities

### New Capabilities

- `pptx-template-conditioned-generation`: Source-bound design-profile selection,
  bounded new-page generation, reimport verification, and package/layout
  evidence.

### Modified Capabilities

- `pptx-design-profile`: Its descriptive evidence becomes an input to a
  documented generation route; it remains read-only and source-bound.
- `pptx-source-derived-reuse`: Repeated source-slide use must cross an
  export/reimport boundary and use ordinal/manifest locators rather than
  position-scoped public IDs.

## Non-Goals

- No universal PPTX AST, raw OOXML editor, HTML/PPTD conversion, or second
  writer.
- No automatic visual similarity claim, model call, image-generation tool, or
  native Windows host requirement in this portable slice.
- No promise that a source deck's pre-existing layout defects are repaired by
  generation; only newly introduced issue categories block the oracle.

## Impact

- Adds a benchmark-side generation runner and fixture test using the public
  OfficeKit API.
- Adds portable Presentation Skill guidance and OpenSpec evidence requirements.
- Real source files remain external inputs identified only by SHA-256.
