# Audience-facing text

Presentation copy is spoken or scanned under time pressure. Preserve facts,
sources, numbers, qualifiers, names, and user-locked wording before editing
style.

Use this editorial sequence:

```text
lock facts and sources
→ state the audience outcome
→ write the page claim
→ add only support needed for that claim
→ fit copy to the rendered page
→ review the whole-deck voice
```

## Write for the page job

- Prefer a conclusion or useful question over a topic label.
- Make the title and dominant visual agree; do not repeat the same sentence in
  title, body, and callout.
- Keep one visible hierarchy: claim, evidence, annotation, source.
- Use direct verbs and concrete nouns. Remove throat-clearing, slogans,
  abstract stacks, and redundant contrast patterns such as repeated
  “not X but Y”.
- Keep uncertainty and scope close to the claim they qualify.
- Put detailed derivation in notes or an appendix only when the visible page
  remains understandable without it.

Simple PPJ text is a string. Mixed formatting uses `paragraphs[]` and `runs[]`.
Do not put Markdown, HTML, CSS, or invented inline markup into a text string.
Assign language and font roles explicitly for mixed-script runs.

For mixed-script text, make the language boundary explicit where shaping,
font fallback, spell checking, or accessibility depends on it:

```json
{
  "runs": [
    { "text": "关键结论", "style": { "language": "zh-CN", "font": "body-cjk" } },
    { "text": " / Evidence", "style": { "language": "en-US", "font": "body-latin" } }
  ]
}
```

Use a bounded BCP-47 tag; do not invent tags from font names or locale display
labels. Language and typeface solve different problems, so setting one does not
authorize guessing the other. Imported direct run language is editable only
through an issued `fontLanguage` leaf.

## Use opacity as hierarchy, not camouflage

PPJ text colors accept eight-digit HEX or a declared color token with `alpha`:

```json
{
  "text": "Source: audited operations ledger",
  "style": {
    "defaultText": { "color": "#16324FB8", "size": 9 }
  }
}
```

The compiler writes one editable native text-color alpha value. Use it for
secondary annotation, metadata, or text over a controlled image overlay. Keep
claims, critical values, and source obligations readable at the intended
delivery distance. Opacity does not repair weak contrast, a busy photograph,
or an unclear hierarchy; adjust the background or composition first.

Imported text opacity may be visible in projected PPJ, but changing source
formatting still requires an issued capability. A text-replacement capability
does not authorize a color or opacity edit.

## Highlight only the evidence that needs it

Use `highlight` on a run when the audience must locate a short phrase, changed
assumption, threshold, or review finding inside otherwise continuous text:

```json
{
  "text": {
    "paragraphs": [{
      "runs": [
        { "text": "Decision: " },
        { "text": "proceed only above 84%", "style": { "highlight": "#FFF2CC" } }
      ]
    }]
  }
}
```

Highlight is editable native text state. It is not a pill, badge, card, or a
substitute for hierarchy. Keep the marked span short and use one highlight
logic consistently across the deck. Authored highlight must be opaque; choose
a lighter color instead of asking alpha to repair contrast.

Imported direct RGB highlights project into `run.style.highlight`. Theme-bound
highlights remain source-owned and are changed only through an issued
`fontHighlightScheme` leaf so PPJ does not flatten a source theme into guessed
RGB.

## Fit without shrinking the argument

Respect the page's content budget. If text does not fit, remove duplication,
split the page by audience task, change the carrier, or move audit detail to a
supporting page. Tiny type is not a valid layout repair.

After every meaningful copy edit, rebuild and render the affected pages. Check
line breaks, widows, CJK punctuation, number/unit binding, source visibility,
and whether the new wording changed the intended visual emphasis.

Local edits stay local. A request to change one title does not authorize a
deck-wide rewrite. Global voice changes require explicit scope and a full-deck
editorial review.
