# Components and templates

Keep these concepts separate:

- A design system is user or brand authority.
- A Presentation Template Skill is style guidance plus representative images.
- A reference PPTX is design evidence or a source package, not automatically a
  template identity.
- A PPJ component is finite reusable program state.

Authority order is: user design system and brand rules, explicitly selected
Template Skill, relevant reference evidence, then self-directed design. Do not
mix multiple unrelated templates.

## Template use

Search and select zero or one template. Read its `SKILL.md` and inspect its
preview and examples. Translate the useful style into the current deck's
Design Grammar: palette roles, typography roles, density rhythm, crop language,
line and surface behavior, carrier choices, motifs, and prohibitions.

Compose pages for the current narrative. Do not reproduce screenshot geometry
pixel for pixel, clone unrelated copy, or force every page into a catalog
layout. A template saves design reasoning; it does not replace it.

A schema-v3 Presentation Template may optionally include `referenceProgram`
and `referencePptx` with exact SHA-256, rights, and source declarations. Use a
reference program only when its license and relevance allow reuse. A reference
PPTX may also be imported as a source-bound artifact; that is a different route
from style guidance.

## Native masters and layouts

Use `design.masters[]`, `design.layouts[]`, and `pages[].layout` when a new deck
needs native PowerPoint layout identity rather than a visual convention only.
This fits a brand title/body system, recurring authoring placeholders, or a
reference PPJ/PPTX that should remain easy to continue in PowerPoint.

The source-free profile is deliberately bounded:

- one canonical master;
- `blank`, `title`, `titleOnly`, or `obj` layouts;
- solid, gradient, image, or none backgrounds already supported by PPJ fills;
- title/body/centered-title/subtitle owner-local placeholders with explicit
  point frames and stable indexes;
- title, body, and other master paragraph defaults at levels zero through
  eight.

Bind every page explicitly when layouts are declared. A page-local
`placeholder` that participates in a layout must use an explicit `index` and
match a master/layout placeholder type and index. PPJ keeps the page-local frame
explicit; it does not pretend inherited native geometry is compiler-owned.

Imported third-party pages expose their stable source layout identity through
`pages[].layout`. Do not invent a new layout ID or add authored master/layout
definitions to a source-bound PPJ: arbitrary imported topology remains in the
source package and such changes fail closed.

## Finite PPJ components

Components may declare parameters, named slots, variants, explicit bindings,
bounded repeat over a finite array, and simple value conditions. Their frame is
local and explicit. Expanded IDs must be deterministic.

An image template element may bind typed component parameters to `image.crop`
or `image.focus` using `{ "crop": { "left": ..., "top": ..., "right": ..., "bottom": ... } }`
or `{ "focus": { "x": ..., "y": ... } }`. A focal point requires `fit: "cover"`
and derives an asymmetric normalized crop; it is not automatic saliency
inference. The compiler keeps both forms bounded and reprojects the same
explicit crop for every repeated instance. Fit, mask, dimensions, and rights
policies remain independently validated by the receiving image slot.

Use components for a repeated semantic structure: a citation row, timeline
event, experimental measure, or brand lockup. Do not turn every page into one
large component or recreate a general layout language inside components.

Forbidden component behavior includes recursion, cyclic dependencies,
unbounded loops, functions, network calls, arbitrary expressions, hidden file
reads, and layout decisions based on executable code. Compiler budgets are
hard limits, not optimization targets.

Review repeated instances for fit and variety. A component may guarantee
structure; it cannot guarantee that each instance communicates well.
