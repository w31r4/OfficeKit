# Source-bound media accessibility metadata

## Why

F-14 already exposes accessibility metadata for recognized shapes, images,
charts, tables, connectors, and groups, but an imported audio/video picture is
projected as opaque and loses its editable `p:cNvPr` accessibility owner. This
leaves an ordinary media alternative-text repair unavailable even when the
media relationship and playback graph can remain untouched.

## What Changes

- Project a canonical imported media `cNvPr` accessibility leaf onto the
  existing common PPJ `accessibility` field.
- Issue `setAccessibility` for media whose residual `cNvPr` profile is
  unambiguous.
- Apply only title, description, and decorative metadata in place while
  preserving media relationships, poster, click action, and timing XML.
- Keep unfamiliar media extension graphs source-owned and fail closed.
- Add one source-bound project → compile → XML → reproject experiment.

## Capabilities

### New Capabilities

- `ppj-source-bound-media-accessibility`: Bounded accessibility metadata on an
  otherwise opaque imported media owner.

### Modified Capabilities

None.

## Impact

- Native PPTX import/projector/compiler and one focused native test.
- Coverage, backlog wording, and media Skill guidance.
- Additive artifact-proto field only; no Office wire-version change. The
  existing `PresentationMedia` accessibility owner and common PPJ element
  field are reused.
- Media payload replacement, playback, trigger/timing, captions, effects,
  and host Accessibility Checker semantics remain outside this slice.
