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

For an ordinary imported text box or shape,
`text.paragraphs[].style.defaultText.bold` and `.italic` edit independent paragraph defaults.
True and false remain explicit; delete the field (or its flag-only
`defaultText` object) to clear the direct default. Restore it by setting the
boolean again. Each field needs its own issued `setTextParagraphStyle` authority;
other defaults, direct run styles and placeholder inheritance are separate.

The same owner supports defaultText.size with its own field authority.
Set a point size from 1 to 768, delete it to clear the direct default,
and set it again to restore it. Native precision is 0.01pt with nearest-even
rounding (18.256 becomes 18.26; 18.125 becomes 18.12); out-of-range sizes are rejected.
Size-only defaultText/paragraph style wrappers can also be removed.
Other paragraph defaults and direct run sizes are preserved.

Use defaultText.fontFamily for the same direct paragraph Latin-font lifecycle,
with its own issued field authority. Names must be nonblank and at most 255
characters. Delete the field or its font-only defaultText/style wrapper to
remove the direct Latin font; set a name again to restore it. Other script
fonts, defaults and direct runs remain unchanged. Extra native font metadata
or child content rejects replacement. This does not prove font availability
or host substitution.

Use defaultText.fontFamilyEastAsia for the independent direct East Asian font.
It has the same name and simple-node limits and supports field or font-only
wrapper deletion and restoration. On source-bound edits, deleting it removes
a:ea even when fontFamily remains; the authored Latin fallback is not reapplied.
Latin/complex-script defaults and direct run fonts are retained.

Use defaultText.fontFamilyComplexScript for independent direct a:cs assignment,
removal and restoration under its exact field authority. It shares the
nonblank/255-character and simple-node limits. Removing a font-only defaultText
or paragraph style wrapper clears it; Latin/East Asian defaults, effects and
direct runs remain unchanged.

Use defaultText.language for independent paragraph-default language assignment,
removal and restoration, including language-only defaultText/style wrappers.
Tags contain 2..63 characters: an initial 2..8 letters, then optional
hyphen-separated 1..8-character alphanumeric subtags (for example en-US or
zh-Hant-TW). Spelling and case are preserved. This is bounded syntax checking,
not a language-registry lookup or proofing-engine validation. Direct run
languages and native altLang remain unchanged. A source lang outside this
grammar stays source-owned: replacing it rejects, while unrelated scalar
edits preserve it.

Use defaultText.kerning for the independent minimum font-size threshold that
enables kerning, in points (0..768). Native precision is 0.01pt with ties-to-even
rounding. Explicit zero remains present; deleting the field or its kerning-only
defaultText/style wrapper removes the direct threshold. Restoration leaves
other defaults and direct run kerning unchanged. Unmodeled native kern remains
source-owned during unrelated edits and rejects replacement. Preview reports
its typography limit; this field does not prove host font shaping.

Use defaultText.letterSpacing for direct character spacing in points
(-768..768). Signed values round to native hundredths with ties to even.
Explicit zero remains present; remove the field or its spacing-only
defaultText/style wrapper to clear the direct spacing, then assign a value to
restore it. Other defaults and direct run spacing stay unchanged. Unmodeled
native spc survives unrelated scalar changes and rejects replacement.
Preview typography remains partial.

Use defaultText.baseline for direct baseline offset in percent (-400..400),
for example baseline: 20 for 20 percent. Native precision is 0.001 percent
with ties-to-even rounding through integer thousandths. Explicit zero remains
present, including tiny negative values rounded to zero. Delete the field or
its baseline-only defaultText/style wrapper to clear it, then assign a value
to restore it. Other defaults and direct run baseline remain unchanged.
Unmodeled source baseline survives unrelated scalar changes and rejects
replacement; host text layout remains unverified.


`textWarpPreset`, `textWarpAdjustments`, `flatTextZ`, `fromWordArt`, `compatibleLineSpacing`, `spaceFirstLastParagraph`, `forceAntiAlias`, `anchorCenter`, `autoFit`, `normalAutoFit`, `margins`, `columns`, `columnGap`, `verticalAlignment`, `upright`, `rotation`, `columnDirection`, `verticalText`, `wrap`, `horizontalOverflow` and `verticalOverflow` belong to the text body: `text.style` on structured table-cell text,
`style` on text and supported owner-local placeholders, and `textStyle` on shapes.
Explicit `true`/`false` for upright and signed degrees/zero for rotation retain
direct native values. Column direction uses `left-to-right` (explicit native false)
and `right-to-left` (true); source removal deletes `rtlCol`. Body rotation is separate from the element's frame rotation.
Text direction uses `horizontal`, `vertical` or `vertical270`; explicit horizontal
retains native `vert=horz`, while source removal deletes `vert`.
Wrapping uses `square` or `none`; explicit none disables wrapping, while source
removal deletes the direct `wrap` override.
Horizontal overflow uses `overflow` or `clip`; source removal deletes `horzOverflow`
without changing `verticalOverflow`.
Vertical overflow uses `overflow`, `ellipsis` or `clip`; source removal deletes
`vertOverflow` without changing `horizontalOverflow`.
Vertical alignment uses `top`, `middle` or `bottom`; middle maps to native center.
Source removal deletes `anchor` without changing the independent `anchorCenter`.
Column spacing uses `columnGap` in points (0..10000). Explicit zero remains a
direct value; source removal deletes `spcCol` and preserves columns/direction.
Column count uses `columns` (integers 1..16). Explicit 1 remains a direct value;
source removal deletes `numCol` and preserves spacing/direction.
Margins use optional `left`, `top`, `right`, `bottom` point values (0..10000).
Zero remains explicit. Delete one edge to remove its direct native inset; empty or
remove `margins` to remove all previously projected insets. Other edges and text remain.
AutoFit uses `none`, `shrink-text` or `resize-shape`. Explicit none writes
`noAutofit`; delete `autoFit` and its dependent `normalAutoFit` to remove the choice.
With shrink-text retained, delete a `normalAutoFit` percentage to remove only that
attribute, or remove the profile to remove both. Explicit fontScale 100 and
lineSpacingReduction 0 remain values. Empty profiles and profiles without
shrink-text remain invalid; preview retains host-reflow limitations.
Anchor-center uses `anchorCenter: true/false`; both retain a direct value.
Delete the projected property to remove `anchorCtr`, independently of
`verticalAlignment`. This deletion requires the updated codec.
The anti-aliasing hint `forceAntiAlias` retains explicit true/false. Delete the
projected property to remove direct `forceAA`, preserving other body state.
Deletion requires the updated codec; the hint does not establish host rendering.
The first/last paragraph-spacing hint `spaceFirstLastParagraph` retains true/false.
Delete it to remove `spcFirstLastPara` with the updated codec while preserving
actual paragraph spacing. Preview retains paragraph-layout limitations.
The compatibility hint `compatibleLineSpacing` retains explicit true/false.
Delete it to remove `compatLnSpc` with the updated codec while preserving
actual paragraph line spacing. Preview retains line-layout limitations.
Text warp uses `textWarpPreset` and optional `textWarpAdjustments` with ordered,
uniquely named signed integer guides. With the preset retained, delete or empty
the list to clear its guides. Delete both fields to remove the canonical warp;
restore them by setting a preset and optional guides. `textNoShape` is an explicit
value. Removal needs the updated codec; unknown warp subgraphs fail closed,
and preview does not render full WordArt.

Flat-text depth `flatTextZ` preserves signed 32-bit integers, including zero.
Deleting the projected field removes canonical `a:flatTx`; restore it by setting
a bounded integer. This needs the updated codec and preserves other source
state. Duplicate or noncanonical children fail closed; preview does not render
full 3D text.

The source marker `fromWordArt` retains explicit true/false. Delete it to remove
the native marker with the updated codec while preserving independent text warp.
The marker does not establish full WordArt geometry or rendering.
On a fresh source PPJ, delete any of these properties to remove that native attribute and
restore inherited/default behavior. You may remove a style object containing only
`textWarpPreset`, `textWarpAdjustments`, `flatTextZ`, `fromWordArt`, `compatibleLineSpacing`, `spaceFirstLastParagraph`, `forceAntiAlias`, `anchorCenter`, `autoFit`, `normalAutoFit`, `margins`, `columns`, `columnGap`, `verticalAlignment`, `upright`, `rotation`, `columnDirection`, `verticalText`, `wrap`, `horizontalOverflow` and/or `verticalOverflow`. If a table cell then projects as plain text, add the
body style through structured text while retaining its native paragraph/run topology.
Keep other body properties and use the issued `setTextBodyStyle` or table-cell
text-style capability; unsupported placeholder owners retain their source boundary.
Preview reports text-layout limits; this field does not prove host rendering.

## Keep the talk track in speaker notes

Use `pages[].notes` for the spoken bridge, caveat, source detail, or facilitation
prompt that a live presenter needs but the audience should not read on the
canvas. Notes use the same `textContent` contract as visible text: a string for
simple prose, or paragraphs and runs when emphasis or language boundaries
matter. Notes supplement the page; they do not excuse an ambiguous visible
claim or hide evidence required for a reader deck.

Imported notes remain source-bound. When the page `nativeRef` issues
`setNotes`, an Agent may change plain text or the text inside an existing rich
run topology. Keep paragraph/run counts, IDs, styles, and representation kind
unchanged. A notes-absent capable page accepts one plain string. NotesMaster,
layout, fields, hyperlinks, picture bullets, arbitrary notes shapes, and
relationship topology remain native-owned and must not be reconstructed.

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

## Use native formulas for mathematical evidence

Keep prose in ordinary runs and put only the mathematical expression in a
formula run. Do not include `\(` or `\)` delimiters:

```json
{
  "runs": [
    { "text": "Expected loss: " },
    {
      "id": "loss-equation",
      "formula": {
        "syntax": "latex",
        "source": "\\sum_{i=1}^{n} p_i \\cdot L_i"
      },
      "style": { "size": 24, "color": "#172033" }
    }
  ]
}
```

The supported subset covers literals, groups, subscript and superscript,
fractions, square roots, common Greek letters and symbols, integral/sum/product,
roman text, common functions, bounded spacing, and ordinary left/right
delimiters. It does not execute TeX: macros, environments, packages,
conditionals, counters, I/O, matrices, alignment, color commands and unknown
commands fail before output. Limits are 4,096 source characters, 512 tokens, 32
nesting levels and 2,048 AST nodes.

A formula run may declare only `size` and solid `color`; it cannot carry a
hyperlink, font override, highlight, gradient, shadow or text decoration.
OfficeKit writes editable native Office Math (`a14:m` with OMML), not an image,
formula font or literal backslash fallback.

Exact LaTeX recovery comes from an OfficeKit-authored PPTX's embedded PPJ.
Third-party OMML stays source-owned and is never reverse-translated into guessed
LaTeX. Do not mutate an imported formula unless a future explicit source-bound
capability says that exact operation is safe.

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

## Reserve gradient and shadow for display text

PPJ can author one native linear or centered-radial text gradient and one
native outer shadow on a run or `defaultText` style:

```json
{
  "text": "Annual signal",
  "style": {
    "defaultText": {
      "size": 34,
      "gradient": {
        "kind": "linear",
        "angle": 18,
        "stops": [
          { "offset": 0, "color": "#16324F" },
          { "offset": 1, "color": "#0B8F8F", "opacity": 0.82 }
        ]
      },
      "shadow": {
        "color": "#16324F66",
        "blur": 3,
        "distance": 1.5,
        "angle": 90
      }
    }
  }
}
```

Use this treatment for a short hero title, launch wordmark, or one display
number when it reinforces the deck's design grammar. Keep body copy, sources,
axes, labels, and evidence text solid. A style cannot declare both `color` and
`gradient`; resolve the paint choice explicitly. Do not use shadow to rescue
poor contrast or place legible text over a busy image. Render the actual slide
and verify the thinnest glyph strokes at delivery size.

The compiler writes editable DrawingML text paint and outer-shadow state, not
a rasterized title. Canonical imported gradients and shadows project back into
PPJ, while theme-transformed gradients, glow, reflection, inner shadow,
WordArt, and irregular effect graphs remain source-owned. A projected effect
does not by itself grant permission to mutate third-party formatting; follow
the issued `nativeRef` capability.

## Style bullets as text, not decoration

Character and numbered bullets may use the same deck-local color tokens and
alpha-bearing literals as other authored text:

```json
{
  "style": {
    "bullet": {
      "type": "character",
      "character": "•",
      "color": { "token": "signal", "alpha": 0.72 },
      "sizePercent": 0.9
    }
  }
}
```

A PPJ palette token resolves to its declared RGB/alpha; it does not silently
claim a native PowerPoint theme identity. Keep bullets subordinate to their
text and use indentation, baseline and spacing for hierarchy. Do not turn each
bullet into a badge, pill or card.

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
