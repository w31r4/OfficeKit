# Design

Authored PPJ design tokens resolve to their declared RGB value, matching runs,
fills and strokes. An alpha-bearing token or literal also sets a presence-aware
paragraph bullet opacity.

The native profile owns only one direct `a:alpha` child under the existing
`a:buClr` RGB or scheme color. Other transforms remain source-owned. Imported
theme colors can project as `{ "token": "accent1", "alpha": 0.5 }`; authored
deck-local tokens compile to direct RGB because a PPJ palette ID is not
necessarily a native DrawingML theme token.
