## Context

An observable PPTD package produced from inline LaTeX writes `a14:m` containing
OMML, not an image. The public Open XML SDK exposes the same Office 2010
Drawing text-math and Office Math vocabulary. OfficeKit can therefore implement
the behavior from the public file format without depending on a browser TeX
engine or a private compiler.

## Decisions

### 1. Formula is an inline run

A PPJ paragraph keeps ordered runs. Exactly one of `text` or `formula` is
present:

```json
{
  "id": "equation",
  "formula": {
    "syntax": "latex",
    "source": "\\int_0^1 x^2 \\, \\mathrm{d}x = \\frac{1}{3}"
  },
  "style": { "size": 28, "color": "#172033" }
}
```

The source excludes `\(` and `\)` delimiters. Formula runs may mix with text
runs in the same paragraph. Only font size and one direct RGB/theme text color
are inherited; hyperlinks, highlights, gradients, shadows, decorations,
letter spacing and arbitrary fonts are rejected for formulas.

### 2. The grammar is finite

The first grammar supports literal identifiers/numbers/punctuation, groups,
subscript, superscript, combined scripts, fractions, square roots, common Greek
letters, common relation/arithmetic symbols, integral/sum/product symbols,
roman text through `\\mathrm`, and spacing commands. `\\left` and `\\right`
may qualify ordinary delimiters.

It does not implement TeX macro definition, environments, packages, file or
network access, conditionals, counters, matrices, alignment, color commands,
arbitrary Unicode control commands or user-defined expansion. Bounds are 4,096
source characters, 512 tokens, nesting depth 32 and 2,048 expanded AST nodes.
Unknown or malformed syntax fails before native output; there is no literal
fallback that silently displays backslashes.

### 3. The wire carries a typed formula AST

The additive wire-v2 content case carries the authored LaTeX plus a finite
recursive math AST. PPJ parsing produces the AST in C#. The PPTX codec writes
OMML from the AST and can read the exact canonical OMML profile back into the
same AST for post-write semantic proof.

The LaTeX string is compiler source provenance, not native identity. Semantic
package comparison masks that source string and compares the AST. Exact PPJ
recovery comes from `/officeKit/program.ppj`, where the original source remains
authoritative.

### 4. Imported formulas remain conservative

Canonical OfficeKit OMML may be inspected internally as formula content.
Ordinary third-party OMML is not translated to LaTeX. PPJ projection exposes a
plain visible summary or opaque/source-bound object and preserves the native
graph. Unsupported mutations fail closed.

## Rejected alternatives

- SVG/PNG formula assets: visually plausible but not native or text-editable.
- Browser-side MathJax/KaTeX: violates direct PPJ-bytes-to-C# compilation and
  adds a second renderer.
- A complete TeX engine: unbounded language surface and unnecessary execution
  power for presentation equations.
- Raw OMML in PPJ: leaks OOXML into the public DSL and makes Agents author XML.

## Lean verification

Extend one existing authored PPJ contract with mixed text/formula runs covering
scripts, fraction, radical, Greek and an integral. Assert native `a14:m`, exact
embedded PPJ recovery, canonical AST re-read and one unsupported-command
rejection. Do not create a formula corpus or renderer snapshot suite.
