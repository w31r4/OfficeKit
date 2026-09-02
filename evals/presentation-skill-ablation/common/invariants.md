# Shared Presentation Invariants

This file is loaded by both experimental Skill arms. It is the control surface:
the arms may differ in route and context packaging, but never in these safety,
fidelity, or visual-quality contracts.

## Start from communication

Before drawing, write:

```text
audience task → one claim → evidence → primary relationship → carrier
→ reading order → protected evidence → canvas occupancy → layer order
```

Use a chart, table, image, diagram, typography, native vector, or deliberate
mix only when it carries the relationship. Empty canvas must create focus,
separation, or rhythm; it must not conceal missing evidence. Do not begin with
`box()`, `card()`, a random shape, or a decorative image.

## Evidence and visual integrity

- Never invent data, sources, cases, citations, or certainty. Mark assumptions
  and placeholders explicitly.
- No card walls, equal rounded panels, pills, badges, colored side strips,
  random circles/rings/arrows, icon clouds, or generic AI rainbow palettes.
- A bar, fill, mask, image, label, number, marker, error bar, line, caption, or
  source may not hide or make another evidence-bearing object unreadable.
- Keep labels outside marks when they collide. Prefer a truthful layout change,
  transparent/offset treatment, or an endpoint comparison over forced overlap.
- With only two observations, use independent endpoint comparison by default;
  never draw a continuous trend through an unknown middle value.
- Do not use tiny text, excessive whitespace, or animation to hide an
  overloaded or incomplete page.

## PPJ and source preservation

- Edit `.ppj` directly through the public `officekit ppj` commands. Do not use
  MJS/JSX, raw OOXML, XPath, relationship IDs, or a second authoring engine.
- Preserve stable element IDs, true back-to-front z-order, source hashes,
  opaque graphs, native references, and non-target package parts.
- Unsupported, stale, ambiguous, or out-of-scope source-bound mutations fail
  closed. Never flatten, rasterize, rebuild, or patch a package to force a
  success.
- Build never overwrites the input. Check, build, render, and review are
  separate evidence stages.

## Images and rights

- An image must carry evidence, identity, explanation, or deliberate atmosphere;
  it is not filler.
- Use the shared `officekit image` route and record the query, candidate,
  selected asset hash, rights, author/license data, crop/focus, and alt text.
- Use user/template/official assets first, then compliant Openverse/Wikimedia
  results or Lucide icons. If no compliant image serves the claim, choose a
  native carrier or state that the asset is missing.
- Images must not cover text, charts, lines, labels, or decision evidence.

## Review and delivery

1. `officekit ppj check`
2. `officekit ppj build`
3. render at final size
4. inspect hierarchy, occupancy, crop, contrast, overflow, and occlusion
5. repair the smallest responsible layer and rerender
6. `officekit ppj review`
7. for imported edits, reimport and audit source/non-target hashes

Structural correctness is not visual approval; rendered evidence is not
PowerPoint playback evidence. State `structural`, `render`, `keynote`, or
`powerpoint` evidence honestly.
