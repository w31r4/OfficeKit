---
name: presentation-template-creator
description: Create or update a reusable PowerPoint style template from a reference deck, visual references, written direction, or an OfficeKit presentation task. Use when the user asks to save, package, contribute, or rebuild a PPT style as a template. Do not use for one-off deck creation, DOCX/XLSX templates, or editing an existing presentation.
---

# Presentation Template Creator

Create OfficeKit's single PowerPoint template form: a complete, self-contained
style Skill, an OfficeKit-authored reference PPTX, and original visual examples.
The main Presentations Skill is a short router; a template Skill is the durable
mini design system an Agent reads after selection. The reference PPTX is a
native, inspectable calibration and asset source for cases where parsing and
source-derived reuse improve fidelity. It is not a fixed layout registry or a
page that every deck must clone. Preserve the useful design evidence an Agent
needs to work without reopening the reference. Do not turn the guide into a
palette slogan, a short list of adjectives, or a compressed copy of the source.
Never copy an external reference deck into a published template: when a clean-
room template needs a PPTX, author a new unrelated deck with OfficeKit and
record its hash and provenance.

Read [references/template-format.md](references/template-format.md) before
packaging a template.

## Workflow

1. Put all references in the current task as read-only inputs. An external PPTX
   is evidence to inspect, not a file to copy into the published template.
2. Render and inspect the references. Distill every reusable, evidence-backed
   decision: audience fit, communication jobs, narrative moves, palette and
   surfaces, typography rhythm, geometry and line language, density, imagery,
   charts, tables, diagrams, motifs, page archetypes, variation limits, and
   anti-patterns. Preserve concrete guidance instead of collapsing it into a
   generic tone word. Unknown facts stay marked as unknown.
   When full-page imagery, scrims, crossing diagrams, foreground labels, or
   other overlaps define the style, also inspect the bottom-to-top scene stack.
   A screenshot proves appearance, not independent editability.
   For image-led references, classify imagery by job before authoring: cover,
   section transition, evidence/context, or atmosphere. A template that uses
   several image moments should carry several original calibration assets (at
   least two when the source visibly changes scene), with deliberate crops or
   placements per role; do not repeat one photograph on every page merely to
   satisfy an image count.
3. Write the style guide independently. It must be useful on its own and
   cover the style's communication territory, visual grammar, page archetypes,
   composition choices, typography and palette roles, content/chart/table/
   diagram treatment, image policy, density/rhythm, layer order, signatures,
   variation limits, prohibitions, review checks, and a calibration map.
   Do not optimize this guide for token count or line count: every published
   template must carry those decisions in a complete, self-contained form. Do
   not copy reference wording or claim uncertain design intent as fact.
4. Create an unrelated four-to-six-page calibration deck with OfficeKit. Use
   new content and geometry, cover at least three page roles, render every page,
   and review it visually. For a template to be called restored, the visual
   fidelity and functional fidelity evidence must each score at least 95/100;
   otherwise publish it as a candidate and record the missing evidence instead
   of rounding the score up. Recreate any design-defining overlap with real
   editable layers, then reopen the PPTX and verify its stack before describing
   the relationship in the template Skill. The reviewed OfficeKit-authored
   deck becomes `assets/reference.pptx`; its pages and native assets are a
   reusable starting point, not mandatory coordinates.
   When imagery is part of the grammar, use role-specific assets in the
   calibration deck and show the treatments the source style actually needs:
   full-bleed, bounded, layered, or cropped. Image choice must carry identity,
   evidence, context, or atmosphere; blank space must not be filled with
   arbitrary stock scenes.
5. Save the complete guide body, the reviewed reference PPTX, calibration PNGs,
   and a packaging spec in the task. External source files, extracted media,
   superseded intermediate decks, and QA evidence stay outside the published
   Skill.
6. Package with:

   ```bash
   officekit run "$SKILL_DIR/scripts/package-presentation-template.mjs" -- \
     --spec <absolute-spec.json> --output-root <absolute-skills-root> --json
   ```

   For an explicit update, also pass the current sidecar SHA-256 as
   `--expected-sha256`.
7. Read the JSON result, query the generated template by exact ID, view its
   preview and examples, and create a short unrelated deck from the style. A
   successful package is not proof of restoration: keep the template pending
   until the reference has been imported, locally edited, re-imported, and
   rendered through the available native host. Stop if the result depends on
   copied layouts, covers an evidence-bearing line, marker, label, connector,
   or image region, or cannot guide free composition.

## Boundaries

- A template supplies style evidence; the Presentations Skill owns production,
  Compose, motion, review, and delivery.
- A user or brand design system remains authoritative. Do not blend conflicting
  templates or select more than one.
- Existing OfficeKit bundled-template migration uses an OfficeKit-authored
  reference PPTX plus recreated examples; it never republishes the source deck
  or old screenshots.
- Do not add `STYLE.md`, source-owned files, executable code, SVG page
  skeletons, fixed-layout registries, or undocumented files to a generated
  template. `assets/reference.pptx` is the one declared native reference asset
  and must be hash-bound in the sidecar.
