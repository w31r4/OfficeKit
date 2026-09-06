# Fonts and typography

Choose typography from audience, language, delivery mode, and design authority.
Do not choose a typeface only because it looks fashionable in a screenshot.

## Define roles

The PPJ `design.fonts` catalog should name a small role system, such as:

- display or cover;
- title and section title;
- body and annotation;
- data and table numerals;
- code or technical notation when required;
- CJK and symbol fallback.

Use `styleRef` for repeated roles and direct overrides only for a deliberate
exception. Preserve the user's brand font when it is available on the target
host. Record fallbacks in the deck's Design Grammar rather than allowing each
text object to improvise.

For an authored theme, `design.theme.fontScheme.majorEastAsia`,
`minorEastAsia`, `majorComplexScript`, and `minorComplexScript` may override
only the corresponding theme slot; when omitted, each falls back to its role
family. These are bounded PPJ owners for the generated theme XML, not a claim
about imported theme editing, font embedding, or host-level font substitution.

## Match the script and medium

- For Chinese or mixed CJK/Latin text, choose families with complete glyph
  coverage, compatible metrics, and appropriate punctuation behavior.
- For multilingual runs, set language and font roles per run when needed.
- For dense numbers, use stable numeral width and verify decimal alignment.
- For technical notation, check symbols, superscripts, subscripts, and code
  glyphs in the final renderer.
- A live deck generally needs larger visible type than a reader handout; do not
  solve reader density by shrinking a live presentation.

Font names are intent, not proof. Renderer substitution can change line breaks,
paragraph height, chart labels, and page balance. Inspect the target host or
render evidence before claiming fidelity.

## Fit and hierarchy

Create contrast with size, weight, position, and whitespace before adding
containers. Keep title, evidence, annotation, and source roles visibly
different. Avoid excessive weights and one-off sizes that destroy rhythm.

If text overflows, edit the copy, adjust the frame, split the page, or change
the carrier. Do not reduce important text below the reading conditions implied
by the delivery mode.

After font or fallback changes, rebuild and render all affected pages. Check
CJK line breaking, punctuation, widows, clipped glyphs, number/unit binding,
text-image contrast, and whether substituted metrics now obscure another
object.
