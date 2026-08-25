## Context

The shipped Presentations Skill already uses an evidence lock, audience rewrite,
page-fit compression, and deck-voice review. Those 42 lines protect meaning but
do not contain the stronger sentence-level patterns in `Trim CN Tech Doc`.

The general trim Skill cannot be reused unchanged. It assumes a human is
reviewing one document section at a time before edits land. Presentation work
often begins from a one-line request, creates all copy before a page exists, and
then revises it after rendering. Slide titles, chart labels, sources, visible
body copy, and speaker notes also have different information budgets.

The authoring plan already has an open `editorial` object and immutable plan
revisions. The task system already binds a reviewed artifact to the current plan.
This change can therefore remain a Skill-level specialization without adding a
runtime text model, another DSL, or another public API.

## Goals / Non-Goals

**Goals:**

- Preserve the directness of `Trim CN Tech Doc` while adapting it to slide
  hierarchy, delivery mode, and page geometry.
- Give the main Presentations workflow one small editorial authority that works
  for creation, template-conditioned creation, and bounded existing-deck edits.
- Keep facts, citations, uncertainty, technical terms, user wording, and local
  edit scope intact.
- Make the full deck sound like one human author without forcing every title
  into the same rhetorical form.

**Non-Goals:**

- Detect whether text was written by AI or produce an "AI score."
- Rewrite facts, infer missing evidence, translate source quotations, or alter
  code and literal source material.
- Replace presentation strategy, narrative planning, layout, or visual review.
- Turn every slide into minimal copy; reader decks may legitimately be dense.
- Add a text-generation model, network provider, JavaScript API, or codec field.

## Decisions

### Add one focused sibling Skill

`presentation-editorial-trim` will live in the existing Presentations plugin and
remain below the domain-level `presentations` route. The main Skill invokes it
at the editorial stages, while a user can call it directly for a copy-only deck
revision.

This is preferable to expanding the already routed Presentations entry with a
long pattern catalogue. It also avoids a dependency on a separately installed
general trim Skill. A second runtime or object model would add no value.

The current `references/audience-text-editing.md` will cease to be a parallel
authority. Its useful four-pass contract moves into the focused Skill or becomes
a short pointer to it.

### Edit four copy layers differently

The Skill classifies text before rewriting:

| Layer | Primary responsibility | Default edit behavior |
|---|---|---|
| Claim/title | Advance the argument or orient the audience | Direct, speakable, varied syntax |
| Visible support | Explain evidence the visual cannot carry | Compress after composition |
| Labels/sources | Identify data, units, provenance, and objects | Preserve exact facts and terminology |
| Speaker notes | Carry delivery support and necessary detail | Keep complete prose appropriate to the presenter |

Brand or section pages may use intentional fragments. Analysis, academic, and
management pages usually need complete claims. This avoids the two bad extremes:
turning every title into a slogan or forcing every title into a consulting-style
declarative sentence.

### Use two passes around composition

The pre-composition pass settles the claim, removes filler, records protected
facts, and gives each text block a role. The post-render pass responds to actual
wrapping and visual hierarchy. It edits copy, splits a page, or changes the
composition before reducing type size.

The second pass cannot change evidence, citations, names, numbers, or the page's
reader task. If a shorter statement changes certainty or causality, the Skill
keeps the longer version and changes the page.

### Adapt patterns as judgments, not regular expressions

The presentation rule set includes false contrast, defensive negation,
throat-clearing, empty report signposts, abstract noun chains, source-code jargon
in audience copy, repeated three-part phrasing, unsupported superlatives,
continuous metaphors, slogan fragments, and repeated title openings.

These are prompts for editorial judgment. A real comparison may use contrast; a
quoted sentence may contain a long dash; a three-step process may have exactly
three steps. The Skill checks whether the form carries information before
changing it. Deterministic review may flag repetition, but it never rewrites copy
or blocks publication on linguistic taste alone.

### Make delivery mode change where information lives

- `live`: visible copy stays sparse enough to speak over; supporting explanation
  moves to notes without removing evidence or sources.
- `reader`: visible copy remains self-contained and retains qualifiers needed for
  independent reading.
- `hybrid`: the claim and evidence interpretation remain visible; procedural talk
  track and optional elaboration move to notes.

This is a placement decision, not permission to discard information.

### Keep edits source- and scope-bound

For imported decks, copy-only polishing uses resolved text nodes and the existing
source-bound edit path. It changes only declared pages and produces a concise list
of edited text roles. Opaque or unsupported text remains unchanged. A request to
change deck-wide voice first updates the authoring plan and declares global scope.

## Risks / Trade-offs

- [Rules remove a deliberate rhetorical contrast] -> Treat patterns as
  candidates, preserve evidenced comparisons, and inspect title variation across
  the whole deck rather than banning one string form.
- [Compression removes a necessary qualifier] -> Lock facts, uncertainty, units,
  dates, names, and citations before editing; prefer a new composition over a
  changed claim.
- [A new sibling Skill increases routing surface] -> Keep it under the
  Presentations plugin, give it one narrow purpose, and let the domain Skill own
  automatic invocation.
- [The external trim Skill and the OfficeKit copy drift apart] -> Port only the
  stable language patterns OfficeKit needs and maintain presentation-specific
  examples locally; do not attempt cross-repository runtime synchronization.
- [Every title becomes terse and monotonous] -> Preserve a small title spectrum:
  claim, question, section marker, instruction, quotation, and intentional
  fragment, selected by page job.

## Migration Plan

1. Add the sibling Skill and its presentation-specific pattern reference.
2. Move the existing four-pass presentation contract into that authority and
   convert the old reference into a short route or remove it in the same change.
3. Wire the five Presentations task routes to the pre-composition, post-render,
   or local-edit pass they need.
4. Update plugin manifests, package inventory, and minimal Skill checks.
5. Exercise the Skill on one new deck and one bounded existing-deck copy edit;
   retain only failures that justify a small regression example.

Rollback removes the sibling Skill and restores the current four-pass reference;
no task, plan, PPTX, API, or wire migration is required.

## Open Questions

- Whether the public Skill name should be `presentation-editorial-trim` or the
  more user-facing `presentation-copy-editor`; the capability contract is the
  same either way.
