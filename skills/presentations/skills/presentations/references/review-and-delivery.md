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
   evidence.
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

For `live` or `hybrid` delivery, also read the speaker notes in sequence. Verify
that the talk track agrees with the visible claim, preserves qualifications and
sources, and does not contain instructions meant for the audience. After a
source-bound notes edit, re-import and confirm the intended page notes while
keeping every non-target slide and native notes graph stable.

Treat `pages[].hidden` as ordinary slide-show routing, not deletion. A hidden
appendix page still belongs to sections, custom shows, review, and delivery and
must remain factually and visually valid. On imported pages, change `name` or
`hidden` only through the page nativeRef capability; set `hidden: false` to show
a page again rather than removing the field. Neither edit changes custom-show
membership or presentation order.

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
