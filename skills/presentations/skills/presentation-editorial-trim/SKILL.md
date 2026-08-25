---
name: "presentation-editorial-trim"
description: "Edit presentation titles, visible copy, labels, sources, and speaker notes into concise audience-facing language while preserving facts and slide scope. Use for PPT/PPTX copy polish, removing AI-like phrasing from slides, or the editorial passes required by the Presentations Skill. Do not use as a general prose editor."
---

# Presentation Editorial Trim

Edit presentation copy without changing what the evidence says or redesigning
the deck. This Skill is the editorial authority for OfficeKit presentations.
The Presentations Skill invokes it during creation and bounded edits; users may
also invoke it directly for copy-only work.

For resumable creation or imported-deck work, keep the copy rules in the active
authoring plan and continue through `officekit repl`. A narrow one-pass copy
review may run directly through the Presentations workflow.

## Establish the boundary

Before rewriting, record:

- the intended audience and delivery mode: `live`, `reader`, or `hybrid`;
- whether the scope is the full deck or named pages only;
- names, numbers, dates, units, quotations, citations, uncertainty, technical
  identifiers, and user-approved wording that must remain exact;
- the current deck voice when editing an existing presentation.

Do not alter a factual value, source meaning, confidence level, or protected
wording to make a sentence shorter. Do not touch non-target pages during a
local edit. If the user asks only for copy editing, preserve layout, object
identity, theme, images, charts, and opaque native content.

## Separate four copy layers

Classify text before editing:

1. **Titles and claims** — the page's conclusion, question, or reader task.
2. **Visible support** — evidence interpretation and the minimum explanation
   the page needs.
3. **Labels and sources** — axes, legends, object labels, citations, units,
   dates, and provenance.
4. **Speaker notes** — talk track, nuance, transitions, and details that help a
   presenter but need not occupy the canvas.

Never solve a crowded slide by deleting labels or source qualifiers. Move
presenter-only detail to notes when the delivery mode allows it.

## Apply two passes

### Pass 1: shape copy before composition

- State the page's one audience-facing point.
- Write the title as a natural claim, useful question, or explicit reader task.
- Attach the supporting evidence and preserve its qualifier.
- Remove content that belongs on another page or in notes.
- Set a visible-copy budget before the page is laid out.

### Pass 2: edit against the rendered page

- Read the title at contact-sheet scale and the support at full-page scale.
- Shorten repetition before shrinking type.
- Repair title, support, and label hierarchy together; do not polish one text
  box in isolation.
- Keep citations and units adjacent to the evidence they qualify.
- Re-render after material copy changes and report the visual-review state.

For an imported deck, limit both passes to the declared pages and preserve the
existing voice unless a deck-wide rewrite was explicitly requested.

## Adapt to delivery mode

- **Live:** keep the canvas sparse enough to scan while speaking. Put nuance,
  transitions, and optional detail in notes.
- **Reader:** make visible copy self-contained. Do not hide necessary reasoning
  in notes that the reader may never see.
- **Hybrid:** keep the claim and evidence interpretation visible; place the
  presenter-specific expansion in notes.

## Trim with judgment

Read [the pattern reference](references/patterns.md) when rewriting prose.
Patterns are review candidates, not banned tokens. A contrast, list, repeated
term, or metaphor may be correct when it expresses the source or the user's
voice. Rewrite only when the pattern weakens meaning, rhythm, or credibility.

Titles across a deck should vary naturally among claims, questions, decisions,
and section turns. Do not force every title into the same grammar. Labels stay
literal; sources stay traceable; notes may sound more conversational than the
visible slide.

## Verify and return

After editing:

- compare all locked facts and source labels with the input;
- read titles in sequence for logic and repeated openings;
- verify the changed page IDs and non-target stability;
- inspect overflow, wrapping, and text/container hierarchy;
- reopen the exported PPTX and render the changed pages when available.

Return the output path and SHA-256, changed pages, preserved fact/source
evidence, any unresolved wording, and `visualReview` as `complete`,
`requires-human`, or `unavailable`.
