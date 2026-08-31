# Reference reconstruction

Use this route when the template source is an existing PPTX, a long montage, a
set of screenshots, or a design guide with visual examples. The goal is not to
trace pages. It is to recover a reusable visual grammar and prove that
OfficeKit can express it with independent, editable content.

## Classify the source

- **Reference PPTX:** render every page, inspect its theme, page roles, object
  classes, z-order, imagery, charts, tables, diagrams, and imported capability
  boundaries. Keep the source read-only and private.
- **Guide plus images:** read the complete guide, inspect the full montage, then
  inspect representative pages at full resolution. The guide explains intent;
  the images remain the appearance oracle.
- **Images only:** infer cautiously. Record repeated observable decisions and
  label unproven rationale as an interpretation.

Do not crop a source montage into template examples. Do not reuse its text,
data, screenshots, photographs, logos, page geometry, or branded composition.

## Build a reconstruction map

Choose four to six representative page roles rather than the prettiest pages.
For each role record:

- communication job and evidence carrier;
- surface and layer stack from back to front;
- type hierarchy and alignment system;
- image role, crop, mask, scrim, and contrast treatment;
- chart, table, diagram, or native geometry behavior;
- density, whitespace purpose, and recurring motif;
- what must remain editable in the clean-room reference;
- which PPJ/compiler capability will express it.

The map should cover the style's difficult case. A photo-led system must include
one real image-backed page. A data-led system must include a native chart or
table. A layered system must include a verified overlap. Do not select only
easy text pages.

## Probe before composing

For every design-defining carrier, make the smallest PPJ that can prove it:

```text
minimal PPJ
→ officekit ppj check
→ officekit ppj build through NativeAOT
→ officekit ppj render through the available native host
→ inspect the rendered pixels
→ re-import or inspect the PPTX when identity/order matters
```

Probe examples include native table text, a full-bleed background with scrim
and foreground text, an image mask, a chart with labels, or a crossing
connector. If the probe fails, treat that as a compiler/runtime defect or an
explicit capability boundary. Fix it before relying on the carrier, or mark
the template candidate incomplete. Do not silently swap in a visually weaker
construction.

## Reconstruct with unrelated content

Create one coherent calibration story with different subject matter, facts,
imagery, and page geometry. Preserve the recovered design logic, not the source
page coordinates. Use distinct original or licensed images for distinct visual
roles. Keep text, charts, tables, shapes, lines, and foreground labels native
and editable wherever the style depends on them.

The reference PPJ is the semantic source. The reference PPTX must be compiled
from it. The PNG examples must be rendered from that PPTX; they are not separate
mockups.

## Score without rounding

Score visual and functional fidelity separately.

Visual fidelity covers hierarchy, silhouette, density, color/surface behavior,
typography rhythm, imagery, chart/table treatment, and layer effects.

Functional fidelity covers native editability, z-order, image/background
behavior, semantic charts/tables/diagrams, stable IDs, re-import, and the absence
of hidden raster substitutions.

Both must be at least 95/100 before calling the reconstruction complete. Record
the weakest role and the missing capability. A high visual score cannot cancel
a flattened or non-editable implementation, and a structurally valid PPTX
cannot cancel a visibly degraded host render.

## Compare and revise

Compare the rendered calibration pages with the source at the level of design
decisions, not pixel identity. Revise the guide only after observing what the
Agent misunderstood or what the compiler could not express. Add a rule when it
prevents a repeated failure; do not turn one source page into a fixed layout.
