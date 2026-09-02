# Creative direction

Use this reference for image-led, brand, product-launch, editorial, or
"make it feel designed" work. It is a positive art-direction layer between
the narrative and PPJ composition. It does not choose a fixed template, copy a
reference page, or replace the visual-attention and evidence contracts.

## Contents

- [Route](#route)
- [Visual thesis](#1-state-the-visual-thesis)
- [Asset board](#2-build-an-asset-board)
- [Page archetype](#3-choose-a-page-archetype)
- [Deck grammar](#4-write-the-deck-grammar)
- [Layer stack](#5-compose-the-visual-stack)
- [Motion](#6-add-motion-only-after-the-page-works-still)
- [Review](#review-questions)

## Route

```text
communication job
→ visual thesis
→ asset board
→ page archetype
→ design grammar
→ layer stack
→ PPJ composition
→ optional motion
→ rendered review
```

Load it when the primary scenario is `brand-creative`, when the page is
image-led, or when the user asks for a visual concept, launch feel, editorial
finish, or stronger visual impact. Keep the selected scenario and design
system authoritative. A template or reference deck supplies evidence about a
style; it does not turn into a page skeleton.

## 1. State the visual thesis

Write one sentence that explains what the page should make the audience feel
or notice, for example: “A machine built for motion feels precise, quiet, and
larger than its enclosure.” The thesis is not a slogan. It determines the
image subject, crop, type scale, surface contrast, and one dominant visual
action. If the thesis cannot be stated, do not start drawing.

## 2. Build an asset board

Before composing an image-led page, name the media roles rather than asking
for “some images”. A small board is enough:

| role | subject / treatment | target pages | rights / status |
| --- | --- | --- | --- |
| hero | the subject that carries the promise; deliberate crop | opening or reveal | user, official, or registered search receipt |
| detail | material, interface, texture, or mechanism | proof/detail | same source family, or omit |
| context | environment that gives scale or use | transition / proof | source and alt text required |
| mark | one small icon or native symbol | navigation / annotation | Lucide or approved brand asset |

Search for the role and subject in short English terms. Register the chosen
candidate with `officekit image add` before putting it in PPJ. Keep the asset
ID, exact hash, crop intent, alt text, rights, and credit line. Use two or
three distinct assets when the narrative benefits from a sequence; do not
repeat one stock crop to make a deck look full. When no compliant image earns
space, use a chart, diagram, native vector, or typography instead.

## 3. Choose a page archetype

Select the page silhouette before choosing coordinates. Vary archetypes across
the deck so the visual rhythm is intentional:

- **Hero reveal** — full-bleed or dominant image, one short claim, one quiet
  piece of furniture.
- **Statement over field** — a photograph or texture establishes atmosphere;
  typography carries the argument with a controlled scrim.
- **Split proof** — one image and one evidence carrier share a baseline; each
  has a different job, not two decorative panels.
- **Detail crop** — a close crop or masked image reveals a material, mechanism,
  or interaction and is annotated by a small number of labels.
- **Image sequence** — two to four different crops form a progression such as
  context → detail → consequence; captions explain the relationship.
- **Type-led release** — no image is stronger when the sentence or number is
  the subject; add only a purposeful line, field, or mark.

These are choices, not required page types. A page may be sparse when the
visual thesis is a deliberate pause, but never to hide missing evidence.

## 4. Write the deck grammar

Record a short, deck-specific grammar in `design.grammar` before composing:

- **surface** — background field, image treatment, scrim strength, and where
  contrast must be protected;
- **type** — display, body, data, and CJK fallback roles plus the scale jump
  between them;
- **image** — subject distance, crop bias, focal-point rule, color treatment,
  and how many distinct assets the story needs;
- **geometry** — one recurring action (diagonal field, frame, rail, wedge,
  crop window, or baseline), with its information role;
- **furniture** — kicker, folio, source, or rule that stays quiet and
  consistent;
- **variation** — which pages are reveal, evidence, detail, transition, or
  conclusion pages and how their density changes;
- **forbidden patterns** — filler stock, random geometry, unlicensed media,
  unreadable type, and any object that hides evidence.

Use two or three palette roles rather than a theme assembled from unrelated
accents. A color field or gradient is a surface device, not a substitute for
an image. Let one large type decision and one spatial decision do most of the
work; do not add a second motif merely because the page has room.

## 5. Compose the visual stack

For an image-led page, keep the order explicit in `pages[].elements[]`:

```text
native image background or image element
→ crop / color field
→ scrim or bounded mask
→ image subject / evidence
→ title and annotation
→ foreground rule, action, or source
```

Use a native background for a true full-page surface. Use an image element
when the image must be selectable, independently cropped, masked, animated,
or moved. The subject may overlap a frame or field when the overlap creates
meaning; it may not cross a title, label, chart mark, or source line. Reserve
clearance around evidence before adding any foreground treatment.

Positive visual effects that remain editable include full-bleed photography
with a restrained scrim, a diagonal or wedge-shaped color field, a large
cropped subject crossing a safe boundary, a thin baseline or rule, a masked
detail window, and one small source/folio system. These effects are optional
and must explain the thesis; they are not a list to exhaust.

## 6. Add motion only after the page works still

The static page must already communicate at final size. Then choose at most
one motion purpose: reveal the evidence, sequence a causal relationship,
focus a key value, or maintain calm continuity. Use the motion reference for
typed effects and triggers. Never use animation to hide an empty composition,
an unreadable crop, or an occluded chart.

## Review questions

Before delivery, render the page and ask:

1. Can I name the page's visual thesis and primary carrier in one breath?
2. Does the chosen image add evidence, identity, explanation, or atmosphere?
3. Is the canvas occupied by the carrier rather than by a row of generic
   containers? If it is sparse, is the pause intentional and recorded?
4. Are title, evidence, subject, labels, and sources readable without one
   covering another at final size?
5. Does this page have a different silhouette or density from its neighbors
   for a narrative reason?
6. Would removing one effect make the claim clearer? If so, remove it.
7. Is every external image registered, hash-bound, rights-labelled, and
   accompanied by the required attribution or human-review note?

Visual richness is a consequence of a coherent thesis, asset sequence, type
scale, surface hierarchy, and spatial contrast. It is not the number of
shapes, colors, or images on the canvas.
