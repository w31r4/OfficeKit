---
name: presentation-template-creator
description: Create or update a reusable PowerPoint style template from a reference deck, visual references, written direction, or an OfficeKit presentation task. Use when the user asks to save, package, contribute, or rebuild a PPT style as a template. Do not use for one-off deck creation, DOCX/XLSX templates, or editing an existing presentation.
---

# Presentation Template Creator

Create OfficeKit's single PowerPoint template form: a concise style Skill plus
original visual examples. Never retain a reference PPTX, fixed layout, page
skeleton, source component, or authoring script in the template.

Read [references/template-format.md](references/template-format.md) before
packaging a template.

During calibration, use the Presentations primitive and typography references
(`skills/presentations/skills/presentations/references/primitives.md` and
`.../references/fonts.md`) as shared capability guidance. Record which native
capabilities and render evidence were exercised in the task; publish only
style decisions and original visual examples, never the implementation map.

## Workflow

1. Put all references in the current task as read-only inputs. A PPTX is a
   reference deck, not a template artifact.
2. Render and inspect the references. Distill only reusable decisions: audience
   fit, palette and surfaces, typography rhythm, geometry and line language,
   density, imagery, charts, diagrams, motifs, and anti-patterns.
   When full-page imagery, scrims, crossing diagrams, foreground labels, or
   other overlaps define the style, also inspect the bottom-to-top scene stack.
   A screenshot proves appearance, not independent editability.
3. Write the style guide independently. Do not copy reference wording or claim
   uncertain design intent as fact.
4. Create an unrelated four-to-six-page calibration deck with OfficeKit. Use
   new content and geometry, cover at least three page roles, render every page,
   and review it visually. Recreate any design-defining overlap with real
   editable layers, then reopen the PPTX and verify its stack before describing
   the relationship in the template Skill. For image-led pages, use
   `slide.setNativeBackgroundImage()` for a true native backdrop, or
   `slide.setBackgroundImage()` plus the ordered `slide.elements` stack when the
   image must be movable, cropped, or animated. Use `element.moveBefore()` /
   `moveAfter()` only after checking the current z-order capability; do not
   simulate layer order by rebuilding the whole slide.
5. Save only the guide body, calibration PNGs, and a packaging spec in the task.
   Source files, intermediate PPTX files, extracted media, and QA evidence stay
   outside the published Skill.
6. Package with:

   ```bash
   officekit run "$SKILL_DIR/scripts/package-presentation-template.mjs" -- \
     --spec <absolute-spec.json> --output-root <absolute-skills-root> --json
   ```

   For an explicit update, also pass the current sidecar SHA-256 as
   `--expected-sha256`.
7. Read the JSON result, query the generated template by exact ID, view its
   preview and examples, and create a short unrelated deck from the style. Stop
   if the result depends on copied layouts, covers an evidence-bearing line,
   marker, label, connector, or image region, or cannot guide free composition.

## Boundaries

- A template supplies style evidence; the Presentations Skill owns production,
  Compose, motion, review, and delivery.
- A user or brand design system remains authoritative. Do not blend conflicting
  templates or select more than one.
- Existing OfficeKit bundled-template migration always uses recreated examples,
  never old screenshots.
- The calibration workflow may use the current presentation primitives—native
  background images, cross-type scene ordering, chart/shape/image composition,
  motion, and `presentation.inspect()`—but the published template contains only
  the resulting style guidance and original PNG evidence. A template never
  promises that a third-party source graph is fully editable.
- Do not add `STYLE.md`, a retained reference, executable code, SVG page
  skeletons, or undocumented files to a generated template.
