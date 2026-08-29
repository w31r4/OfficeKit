# Source continuation

Use this route when an imported PPTX remains the actual starting state: the
user wants its pages, components, and native structure reused and then edited.
It is different from a Template Skill or a style-transfer reference, which
provide visual evidence while new pages are freely composed.

## Contract

The source is immutable and authoritative. A continuation transaction must:

1. copy and hash the source in the task workspace;
2. import and inspect every source slide, including renders, layouts,
   placeholders, source assets, and `slide.cloneCapability`;
3. write a complete frame map from every output slide to one source slide;
4. duplicate only capability-supported slides, exporting and reimporting each
   pending clone before starting the next one;
5. re-inspect the final starter IDs, edit only declared inherited targets, then
   export, reimport, and review the result;
6. publish only a distinct output after source, map, manifest, and footprint
   hashes agree.

Closed ownership graphs may retain unknown parts and external relationships.
Unsupported sections, custom shows, jumps, shared descendants, ambiguous
relationships, or graph-budget overflow fail closed before a starter artifact
is written. The explicit `--allow-closed-leaves` option may include one
canonical NotesSlide or legacy comment graph; rich/modern comment graphs still
fail closed. Do not substitute a reconstructed slide or a shared-part copy.

## Frame map and starter

Run the shipped inspection command before planning copy:

```sh
officekit run "$SKILL_DIR/template_following_scripts/inspect_template_deck.mjs" \
  --workspace "$TASK_DIR" --pptx "<source.pptx>"
```

Record every output slide, its source ordinal, narrative role, inherited edit
targets, and any omitted source slide with a reason:

```json
{
  "outputSlides": [
    {"outputSlide": 1, "sourceSlide": 3,
     "narrativeRole": "opening thesis", "reuseMode": "duplicate-slide",
     "editTargets": []}
  ],
  "omittedSourceSlides": [
    {"sourceSlide": 4, "reason": "appendix not needed"}
  ]
}
```

Validate the map, then use
`prepare_template_starter_deck.mjs`. The command independently revalidates the
immutable source, performs one clone per export/reimport boundary, translates
source locators to fresh starter IDs, and preflights deletion of omitted source
slides. It writes no PPTX, preview, layout, contact sheet, or manifest when a
source graph cannot be proved safe.

The starter manifest is the only locator authority after reimport. It records
source/output slide pairs, source and starter element IDs, inherited locators,
source/map/inspection/output hashes, clone boundaries, deletion boundary,
provider, and no-overwrite policy. IDs from the inspection NDJSON are not
assumed to persist.

## Editing the continuation

Use the starter as the authoring base and bind one
`office-kit.template-edit-plan.v1` to its PPTX and manifest hashes. Every
manifest target is listed exactly once, including `operations: []` for a kept
target. Use only the bounded typed operations exposed by the imported
capabilities index: fixed-topology text, position, table cell, chart title/data,
same-format image replacement, and capability-proven top-level deletion.

Inherited typography, placeholders, crop, accessibility metadata, native IDs,
and layout bindings remain source-owned. Fill or remove visible inherited
placeholders intentionally; never leave prompt text such as `Slide Number`,
`Date`, or `Footer` in the output. If content does not fit, shorten it, choose
another inspected source slide, or split the page. Do not silently shrink or
redraw the source design.

After every source-bound export, reimport and resolve fresh IDs. Review against
the immutable source with exact changed pages and package footprint. A local
continuation edit must not normalize untouched pages or introduce arbitrary
tables, charts, connectors, groups, native nodes, or raw XML. If the requested
target is outside the issued capability, report the boundary and publish no
partial artifact.

Use the existing scripts for inspection, starter creation, plan validation, and
fidelity checks; they are executable transactions, not visual templates. Keep
their temporary outputs under the task directory and leave the input untouched.
