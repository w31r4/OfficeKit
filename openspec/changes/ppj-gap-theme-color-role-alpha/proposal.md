# Proposal: Preserve authored theme-role alpha

Add a small authored-only theme color profile for alpha-bearing accent and
dark/light/hyperlink roles. Existing six-digit values remain valid; the new
`#RRGGBBAA` spelling lowers to the corresponding DrawingML `a:alpha` child
under `a:srgbClr`.

This closes one concrete part of the F-15 theme-transform gap without
claiming a general imported-theme editor. Source-bound programs that declare
the authored theme owners continue to fail closed, and tint/shade/effect
scheme transforms remain outside this slice.
