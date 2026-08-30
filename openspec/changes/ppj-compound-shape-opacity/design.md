# Design

`shape.style.opacity` is a semantic multiplier from zero through one. For a
source-free shape, the compiler applies it after resolving named and inline
styles and after constructing the text body:

```text
effective branch alpha = branch-local alpha × shape opacity
```

The bounded branches are direct solid/gradient/image fill, visible outline,
outer shadow, explicit run/default-run solid or gradient text paint, text
shadow and explicit bullet color. Alpha is written into the existing wire-v2
fields before the canonical PPTX writer runs, so no protocol change is needed.

An opacity of one is a no-op. A lower opacity fails before output if visible
text relies on an unresolved inherited color, or if an explicit highlight is
present because DrawingML highlight has no bounded alpha surface. This is more
accurate than resolving theme inheritance to an arbitrary RGB value.
