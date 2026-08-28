## 1. Freeze the contract and evidence

- [x] 1.1 Reuse the three external PPTX hashes and package inventories from the lossless benchmark.
- [ ] 1.2 Record current direct source order, element kinds, dependency risks, and reorder capability for each sample.
- [x] 1.3 Add one source-free photo/scrim/text reference page that reproduces the fixed-bucket failure.

## 2. Implement the authored scene stack

- [x] 2.1 Add ordered `slide.elements` while retaining type collections as filtered indexes.
- [x] 2.2 Register add/delete/group operations with exactly one owner stack.
- [x] 2.3 Add shared ordering capability and `sendToBack`, `bringToFront`, `moveBefore`, `moveAfter` methods.
- [x] 2.4 Export, SVG preview, layout inspection, and model serialization in scene-stack order.
- [ ] 2.5 Add `slide.setBackgroundImage(...)` and prove image -> scrim -> editable foreground output.

## 3. Preserve and edit imported order

- [x] 3.1 Hydrate source order without type regrouping and expose stack index plus source revision.
- [x] 3.2 Issue a bounded direct-element reorder capability with explicit blocked reasons.
- [x] 3.3 Apply proven reorder in the C# codec by moving existing native nodes; reject unsafe dependencies and mixed mutations.
- [ ] 3.4 Reimport and continue editing without changing unrelated OPC parts, relationships, or opaque objects.

## 4. Agent surface

- [ ] 4.1 Add Help for scene-stack inspection, ordering operations, and background images.
- [ ] 4.2 Update Presentations Skill with image-led composition, scrim/contrast, obstruction, and layer review rules.
- [ ] 4.3 Update Presentation Template Creator so image-led examples demonstrate real layer composition and remain optional where inappropriate.

## 5. Acceptance

- [x] 5.1 Run one compact public-contract test for cross-type authored order and background-image round trip.
- [ ] 5.2 Run imported inspect/reorder/preserve/reimport checks on every sample; mutate only capability-proven targets.
- [ ] 5.3 Render and visually inspect the photo/scrim/text reference and supported imported edits in the available host.
- [ ] 5.4 Repeat from a packed clean install and one fresh Agent context.
- [ ] 5.5 Run final required full/package gates once and publish the verified and blocked boundaries.
