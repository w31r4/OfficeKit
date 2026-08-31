---
name: presentation-template-creator
description: Create or update a reusable PowerPoint style template from a reference deck, visual references, written direction, or an OfficeKit presentation task. Use when the user asks to save, package, contribute, or rebuild a PPT style as a template. Do not use for one-off deck creation, DOCX/XLSX templates, or editing an existing presentation.
---

# Presentation Template Creator

Create OfficeKit's single PowerPoint template form: a complete, self-contained
style Skill plus original visual examples, with an optional clean-room
reference PPJ/PPTX. The guide must preserve the design decisions an Agent needs
without reopening the source; do not compress it into palette names and a few
adjectives.
Never publish the user's source deck, a copied layout, an executable authoring
script, or a reference whose rights and provenance are unclear.

Read [references/template-format.md](references/template-format.md) before
packaging a template.
When rebuilding from an existing PPTX, a long reference image, or a design
guide, also read
[references/reference-reconstruction.md](references/reference-reconstruction.md).

During calibration, use the Presentations PPJ, typography, media/layers,
shapes, charts/tables, and components/templates references as shared language
and design guidance. Record which compiler capabilities and render evidence
were exercised in the task; publish only style decisions and original visual
examples unless the schema-v3 template explicitly includes a licensed
reference PPJ/PPTX.

## Workflow

1. Put all inputs in the current task as read-only references. A user PPTX is
   evidence for distillation, never the file published as the template
   reference.
2. Render and inspect the references. Distill only reusable decisions: audience
   fit, palette and surfaces, typography rhythm, geometry and line language,
   density, imagery, charts, diagrams, motifs, and anti-patterns.
   When full-page imagery, scrims, crossing diagrams, foreground labels, or
   other overlaps define the style, also inspect the bottom-to-top scene stack.
   A screenshot proves appearance, not independent editability.
   When imagery is part of the grammar, classify each image by job: identity,
   evidence, context, section transition, or atmosphere. Use distinct original
   calibration assets for distinct recurring roles; do not repeat one image
   across the deck merely to satisfy an image count.
3. Write the style guide independently. Cover communication territory, visual
   grammar, page archetypes, typography and palette roles, image/chart/table/
   diagram treatment, density and rhythm, layer order, signatures, variation
   limits, prohibitions, review checks, and a calibration map. Do not copy
   reference wording or claim uncertain design intent as fact.
4. Before building the calibration deck, run one minimal native capability
   probe for every design-defining carrier the style depends on: for example a
   table, chart, background image plus scrim, masked image, layered diagram, or
   motion recipe. The probe must use the packaged NativeAOT codec and a real
   host render. If the carrier fails, fix or explicitly bound the product gap;
   do not silently replace it with unrelated shapes merely to finish the
   template.
5. Create an unrelated four-to-six-page clean-room calibration deck as PPJ and
   compile it to PPTX. Use
   new content and geometry, cover at least three page roles, render every page,
   and review it visually. Recreate any design-defining overlap with real
   editable layers, then reopen the PPTX and verify its stack before describing
   the relationship in the template Skill. For image-led pages, use
   a page background for a true native backdrop, or an image element in the
   ordered `pages[].elements[]` stack when it must be movable, cropped, or
   animated. Do not simulate layer order by rebuilding the whole slide.
   Score the result against the declared visual and functional fidelity rubric.
   Both scores must reach 95/100 before calling the style restored; otherwise
   label it a candidate and record the missing evidence instead of rounding up.
6. Check, build, render, review, and re-import the clean-room PPJ/PPTX. Inspect
   the actual host-rendered pages, not only PPJ structure or an internal
   preview. A successful build is not evidence that table text, fonts, image
   masks, crop, transparency, or layer order survived the host.
   Keep the
   original source files, extracted media, and analysis evidence private. Add
   the clean-room `referenceProgram` and `referencePptx` to the packaging spec;
   the Creator validates every relative PPJ asset hash and packages those files
   beside the program so the published `reference.ppj` builds standalone
   only when their license, package size, and reuse value justify publishing;
   otherwise publish only the guide and calibration PNGs.
7. Package with:

   ```bash
   officekit run "$SKILL_DIR/scripts/package-presentation-template.mjs" -- \
     --spec <absolute-spec.json> --output-root <absolute-skills-root> --json
   ```

   For an explicit update, also pass the current sidecar SHA-256 as
   `--expected-sha256`.
8. Read the JSON result, query the generated template by exact ID, view its
   preview and examples, and create a short unrelated deck from the style. If a
   reference PPJ/PPTX ships, also import it, perform one bounded local edit,
   re-import, and render it. Stop if the result depends on copied layouts,
   covers an evidence-bearing line, marker, label, connector, or image region,
   or cannot guide free composition.

## Boundaries

- A template supplies style evidence; the Presentations Skill owns PPJ,
  compilation, motion, review, and delivery.
- A user or brand design system remains authoritative. Do not blend conflicting
  templates or select more than one.
- Existing OfficeKit bundled-template migration always uses recreated examples,
  never old screenshots.
- Do not hide a compiler or host-rendering defect by redrawing a semantic table,
  chart, or diagram as loose text and lines. Such a construction is acceptable
  only when it is itself the intended visual grammar and remains accessible and
  editable.
- The calibration workflow may use PPJ background images, true scene order,
  chart/shape/image composition, motion, and `officekit ppj inspect`. A template
  never promises that a third-party source graph is fully editable.
- Do not add `STYLE.md`, an original input deck, executable code, SVG page
  skeletons, or undocumented files to a generated template. Optional reference
  files must be the reviewed clean-room PPJ/PPTX declared in schema v3.
