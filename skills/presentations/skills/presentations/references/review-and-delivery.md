# Review and delivery

Review the artifact in layers. A later layer does not replace an earlier one.

1. **Facts**: values, claims, units, names, dates, sources, assumptions, and
   locked text match the supplied evidence.
2. **Communication**: the audience, expected change, delivery mode, and
   after-use are explicit.
3. **Narrative**: page responsibilities form a coherent sequence and the
   requested decision or action follows from the evidence.
4. **Cognitive and editorial**: titles state useful claims, text fits, reading
   order is clear, and detail matches the page budget.
5. **Visual and layer**: carrier, hierarchy, density, crop, contrast, and
   rhythm work at final size; no fill, bar, mask, decoration, or label obscures
   evidence. Compare the rendered page with its visual attention contract:
   reading order, deliberate occupancy, protected evidence, and true layer
   order must agree with the plan.
6. **Native and delivery**: PPJ checks, PPTX builds, re-imports, source-bound
   constraints hold, and final paths and hashes are unambiguous.

Run the explicit stages:

```bash
officekit ppj check deck.ppj --json
officekit ppj build deck.ppj -o candidate.pptx --json
officekit ppj render deck.ppj -o previews/ --json
officekit ppj review deck.ppj --json
```

Inspect the rendered pages, not only thumbnails or XML. Review high-risk pages
at full size: dense evidence, combo charts, tables, CJK text, image crops,
layered backgrounds, diagrams, and animated targets. Check non-target pages
after imported edits.

For every chart or layered composition, apply the four executable contracts in
`visual-attention.md`: explicit label budget, honest missing-data topology,
back-to-front evidence-safe layering, and one render-correction loop. If the
rendered page fails any of these, keep the candidate in working state, repair
the smallest PPJ field, and rebuild/re-render. A passing `check` or `review`
report cannot override a visible collision or a chart that implies data that
was not supplied.

For `live` or `hybrid` delivery, also read the speaker notes in sequence. Verify
that the talk track agrees with the visible claim, preserves qualifications and
sources, and does not contain instructions meant for the audience. After a
source-bound notes edit, re-import and confirm the intended page notes while
keeping every non-target slide and native notes graph stable.

Treat an imported `design.canvas` edit as a whole-deck composition change.
Change `width` or `height` only when the canvas nativeRef advertises
`setCanvas`; keep the nativeRef and `unit: "pt"` unchanged. This operation
changes the native page size without scaling, reflowing, cropping, or moving
any page object. Re-import the output and render every page to check exposed
margins, clipping, background coverage, and altered visual balance.

For an imported picture with both `asset` and `svgAsset`, change only
`svgAsset` when its nativeRef advertises `replaceSvg/image.svgAsset`. Re-import
and verify that the same image ID, frame and raster fallback hash survive while
the referenced SVG hash changes. Render the target page in a modern host. Do
not claim legacy fallback parity unless that unchanged raster member was also
tested in the target legacy host.

Treat `pages[].hidden` as ordinary slide-show routing, not deletion. A hidden
appendix page still belongs to sections, custom shows, review, and delivery and
must remain factually and visually valid. On imported pages, change `name` or
`hidden` only through the page nativeRef capability; set `hidden: false` to show
a page again rather than removing the field. Neither edit changes custom-show
membership or presentation order.

Sections and custom shows are different route structures. `sections[]` must
partition every page exactly once in presentation order; `customShows[]` are
named ordered subsets and may repeat a page. On imported decks, change `name`
or `pages` only when that item's nativeRef advertises `setName` or `setPages`.
Keep array count, order, item ID, and nativeRef unchanged, then re-import and
review both the ordinary slide sequence and every alternate show route.

For a capable imported page move, reorder only the existing `pages[]` entries
whose nativeRefs advertise `reorder/pageOrder`. Keep their IDs and nativeRefs
unchanged, do not combine deletion with the move, and update modeled section
membership so it remains a complete partition in the new order. Re-import and
verify the page IDs, every unchanged page-local element ID, comment page
binding, section partition, and custom-show membership. Element
`reorder/zOrder` is a separate in-page operation.

Imported legacy comments are source-bound review evidence. Edit an existing
`comments[].text` only when that comment's nativeRef advertises `replaceText`;
keep its ID, page, author, timestamp, position, resolution state, nativeRef, and
array order unchanged. Re-import and confirm the edited text. Adding, removing,
reordering, resolving, replying to, or changing metadata on imported comments
is outside this bounded profile and must fail closed.

Compilation proves that the program lowered. Re-import proves that OfficeKit
can read the result. Structural motion evidence proves canonical timing state.
None of these proves that a human saw correct rendering or that desktop
PowerPoint played the file without repair.

When no visual-capable reviewer is available, say `visualReview: unavailable`
or `requires-human`; do not infer aesthetic success from layout checks. Record
Keynote or PowerPoint playback separately when actually observed.

Task-bound work commits only the current PPJ revision, candidate, and review
that refer to the same source and program hash. A PPJ change invalidates the
old candidate until it is rebuilt and reviewed.

Deliver the absolute PPJ and PPTX paths, SHA-256, page count, review status,
and useful evidence paths. Preserve the original input. Do not label previews,
scratch builders, temporary assets, or an unreviewed candidate as final.
