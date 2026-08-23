# Template-Conditioned Generation

Use this mode when a user supplies an existing PPTX and asks for new content,
new pages, or a new deck in the same visual system. It is different from
template-following: template-following fills a declared frame map, while this
mode uses the imported deck as a bounded source of page patterns and assets.

## Workflow

1. Copy the input to a task workspace and record its SHA-256. Never edit the
   input in place.
2. Import it with `PresentationFile.importPptx`, then call
   `presentation.designProfile({ maxItems: 64 })`. For multi-page requests,
   pass the user's page roles/content to
   `presentation.planTemplateGeneration({ slides })` and keep its returned
   frame map with the task. Treat the profile as design
   evidence, not as mutation permission. Record the canvas, type scale, palette,
   density, layout families, slide archetypes, component candidates, opaque
   objects, and unresolved/blocked items.
3. Choose zero or one source design system. The plan selects source slides by
   narrative role, content density, preferred visual kinds, and
   `cloneCapability.supported`; prefer distinct archetypes before reusing a
   source slide. Keep the frame map with source ordinal, target role, content,
   bounded edit targets, asset candidates, alternatives, and fit warnings.
4. Duplicate selected source slides through an export/reimport boundary. A
   public slide id is position-scoped and may change after import; use the
   source-bound manifest/ordinal locator, never a stale id. If
   `continuationCapability` is `pending-clone`, export/reimport first. If the
   source graph is not closed and independently owned, report the blocker.
5. Edit inherited text with run-scoped `shape.text.replace`, or use an exposed
   bounded image/SVG-text operation. A reopened slide whose
   `continuationCapability` reports ready `bounded-overlay` may add only its
   listed shapes/images in a clean export; commit/reopen before other SlidePart
   edits. Keep the inherited fonts, geometry,
   paragraph/run topology, placeholders, brand marks, relationships, and
   opaque descendants. If copy does not fit, shorten it or choose another
   source frame; do not silently shrink typography or rebuild the slide.
6. Export and reimport before review. Run `verify`, `validateLayout`, and the
   package oracle. Compare output issues with the source baseline so pre-existing
   source defects are reported as inherited, while new issue categories remain
   blocking. Render every generated slide when the renderer supports it.
7. Complete visual review when visual input is available. Otherwise report
   `visualReview: "unavailable"` (or `requires-human`) and include structural,
   layout, package, and text evidence; never call a montage a substitute for
   visual judgment. A renderer refusal is an explicit capability gap, not a
   reason to flatten custom geometry.
8. Deliver a new output path, source/output hashes, frame map, selected profile,
   verification results, mutation/package footprint, and any inherited or
   unresolved limitations.

## Boundaries

- A template contributes a proven visual grammar; it does not guarantee that
  arbitrary new content will fit every frame.
- Unsupported SmartArt, animation, OLE, modern comments, shared relationships,
  custom geometry, and other opaque graphs stay intact or make the particular
  clone/edit fail closed.
- Do not use raw OOXML, XPath, arbitrary relationship edits, a second writer,
  HTML/PPTD conversion, or a host-specific artifact runtime as a shortcut.
- Do not promise lossless visual fidelity when the source renderer or native
  Office host has not been run. Separate source preservation from visual QA.
