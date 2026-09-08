# Proposal: preserve the authored complex-script font owner

Add one explicit PPJ text-style field, `fontFamilyComplexScript`, for the
DrawingML `a:cs` typeface. The field is available on ordinary text styles and
the bounded PPJ chart text style, survives native import projection, and can
be changed as one revision-bound source-owned run leaf.

This closes the direct complex-script font mapping gap without implementing
host font fallback, glyph shaping, locale inference, or a general theme font
editor. Omitted values remain omitted so existing Latin/East Asian behavior
and inherited source typography stay unchanged.
