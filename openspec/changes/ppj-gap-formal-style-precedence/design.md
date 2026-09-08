## Context

The current compiler already resolves named text styles, inline styles, and
table-specific middle styles. `stylePrecedence` can reorder a few of those
layers, but it cannot name the distinct host layers that matter for text:
layout/master defaults, an element-level text style, a paragraph default run
style, and a direct run style. The JavaScript review path also has no distinct
layout source and currently groups layout and master under `master`.

This change is deliberately authored-only. The effective value is written to
the native run, so a PPTX projection can prove the resulting value, but it
does not claim that PowerPoint's master/layout style graph can be edited from
the new fields.

## Decisions

### Source vocabulary and precedence

`design.grammar.stylePrecedence[].sources` keeps its existing “first declared
source that has a value wins” rule. The formal host hierarchy is therefore
encoded from most specific to least specific:

```text
run → paragraph → element → styleRef → layout → master → theme → default
```

The existing `inline` source remains a compatibility alias that checks the
direct run, direct element text style, and legacy inline default style in that
order. Existing declarations do not change meaning.

The new compiler-aware sources are accepted for the `text.*` targets listed
below. Other existing target families keep their current bounded resolver and
cannot infer layout or paragraph/run owners.

### PPJ owners

- `design.masters[].style` and `design.layouts[].style` are direct
  `textStyle` objects used as fallback text defaults.
- `text` and `placeholder` elements may carry `textStyle` as their direct
  element layer. Their existing `style` field remains the text-container
  (`bodyPr`) style.
- `text.paragraphs[].style.defaultText` is the paragraph layer.
- `text.paragraphs[].runs[].style` is the run layer.
- A named `design.styles.text[].style.defaultText` remains the `styleRef`
  layer.

Only `text.size`, `text.bold`, `text.italic`, `text.font`, and
`text.fontFamily` use this new complete layer map. Missing values continue to
fall through; a `default` source may resolve a declared grammar token.

### Native lowering

The authored compiler resolves each run independently and writes the selected
literal into the existing native run properties. Layout/master styles are not
materialized as new native master/layout style nodes by this slice. Therefore
the experiment verifies effective run values after removing the embedded PPJ
snapshot, not preservation of the declaration graph.

### Review behavior

`reviewPpjArtifact` resolves the same source names for projected/authored PPJ
records and reports the selected source. For a page, layout and master are
resolved from `page.layout`; the review reads their optional `style` objects.
Paragraph and run entries are reported for each rich-text run. Missing owners
are simply absent, not synthesized.

Unsupported source names remain schema-invalid. A declared source may be
present but have no value for a particular element; it is skipped by the
first-hit resolver. Source-bound edits that change the new declaration fields
do not receive a capability in this change and must fail closed.

## Alternatives considered

- Replacing `stylePrecedence` with a global fixed cascade: rejected because it
  would silently change existing declarations that intentionally put
  `styleRef` before `inline` for particular fields.
- Treating `master` as a combined layout/master source forever: rejected
  because it prevents diagnostics and experiments from proving which owner
  supplied a value.
- Writing new native master/layout style nodes: deferred; it requires a
  source-bound OOXML ownership and inheritance proof beyond this field slice.
