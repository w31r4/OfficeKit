## Why

The benchmark already extracts useful design evidence from imported PPTX files,
but that evidence currently lives in a repository evaluator. An Agent using the
public package should be able to ask the imported deck what visual language and
reusable structure it actually contains before planning a continuation.

## What Changes

- Add `presentation.designProfile(options)` as a deterministic, bounded,
  read-only runtime primitive.
- Bind imported profiles to the exact source revision SHA-256; source-free
  profiles remain descriptive and carry no source authority.
- Report canvas, palette, typography, density, normalized geometry rhythm,
  layout families, slide archetypes, repeated visual candidates, and opaque
  native-object summaries.
- Keep the profile free of raw XML, package paths, source bytes, and arbitrary
  selectors. It is evidence for template-conditioned generation, not a new
  universal AST or mutation surface.
- Keep the existing evaluator profile and source-derived reuse operations
  compatible; this runtime slice only exposes a compact public view.

## Non-Goals

- No automatic template selection or model inference.
- No mutation authority, raw OOXML access, or candidate permission bypass.
- No HTML/PPTD conversion and no Windows host acceptance dependency.
