---
name: presentation-template-creator
description: Create or update a reusable PowerPoint style template from a reference deck, visual references, written direction, or an OfficeKit presentation task. Use when the user asks to save, package, contribute, or rebuild a PPT style as a template. Do not use for one-off deck creation, DOCX/XLSX templates, or editing an existing presentation.
---

# Presentation Template Creator

Create OfficeKit's single PowerPoint template form: a complete, self-contained
style Skill plus original visual examples. The main Presentations Skill is a
short router; a template Skill is the durable mini design system an Agent reads
after selection. Preserve the useful design evidence an Agent needs to work
without the reference at hand. Do not turn it into a palette slogan, a short
list of adjectives, or a compressed copy of the source. Never retain a
reference PPTX, fixed layout, page skeleton, source component, or authoring
script in the template.

Read [references/template-format.md](references/template-format.md) before
packaging a template.

## Workflow

1. Put all references in the current task as read-only inputs. A PPTX is a
   reference deck, not a template artifact.
2. Render and inspect the references. Distill every reusable, evidence-backed
   decision: audience fit, communication jobs, narrative moves, palette and
   surfaces, typography rhythm, geometry and line language, density, imagery,
   charts, tables, diagrams, motifs, page archetypes, variation limits, and
   anti-patterns. Preserve concrete guidance instead of collapsing it into a
   generic tone word. Unknown facts stay marked as unknown.
   When full-page imagery, scrims, crossing diagrams, foreground labels, or
   other overlaps define the style, also inspect the bottom-to-top scene stack.
   A screenshot proves appearance, not independent editability.
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
   and review it visually. Recreate any design-defining overlap with real
   editable layers, then reopen the PPTX and verify its stack before describing
   the relationship in the template Skill.
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
- Do not add `STYLE.md`, a retained reference, executable code, SVG page
  skeletons, or undocumented files to a generated template.
