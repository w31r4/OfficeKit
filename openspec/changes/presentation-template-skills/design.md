## Context

Presentation templates are currently distributed as schema-v2 Skills that
retain a reference PPTX, while the Presentations Skill also embeds a fixed Grid
layout library. Search, materialization, tests, package evidence, and authoring
guidance therefore preserve two incompatible authoring models. OfficeKit 1.0
already has free Compose, design grammar, rendering, and review, so templates
only need to condition those capabilities with reusable style evidence.

## Goals / Non-Goals

**Goals:**

- Give `template` exactly one meaning for PowerPoint: a style Skill plus
  original visual examples.
- Make creation, search, selection, task resume, and review consume that form.
- Provide a dedicated creator that can distill a reference without shipping it.
- Rebuild all eight bundled presentation styles and remove the source-backed
  and fixed-layout implementations.

**Non-Goals:**

- Change DOCX/XLSX template semantics, Office wire, presentation codecs, or
  public JavaScript authoring APIs.
- Retain a compatibility path for presentation schema v2 or embedded Grid.
- Copy reference screenshots, layout geometry, code, wording, or private
  implementation into the new first-party styles.
- Rewrite Git history.

## Decisions

### Use schema v3 only for presentation style Skills

`artifact-template.json` schema v3 is valid only with
`kind: "presentation"`. It records the fixed `SKILL.md`, one preview PNG,
four to six example PNGs with roles and hashes, English search metadata, and
provenance. It has no `reference`, `editProfile`, layout registry, or template
kind discriminator. Schema v2 remains valid only for document and spreadsheet
templates. This makes the schema version itself the single format boundary.

### Keep the Skill as the only style instruction source

Each template's `SKILL.md` contains only its distinctive visual grammar,
appropriate uses, and interpretation of its examples. Common presentation
workflow and safety rules stay in the Presentations Skill. No `STYLE.md`, code,
page skeleton, or duplicate instruction document is generated.

### Keep creator reasoning separate from deterministic packaging

The `presentation-template-creator` Skill guides reference inspection, rights
classification, style distillation, and creation of an unrelated calibration
deck. Its script accepts completed guide/metadata/example inputs, validates
them, generates the preview montage and hashes, and atomically publishes the
Skill. Source references and temporary decks remain task evidence and are
never copied into the template directory.

### Preserve search but remove materialization for presentation templates

Search continues to rank local Skills with BM25F and returns
`selectionMade: false`. Presentation candidates expose `skillPath`,
`previewPath`, and `examplePaths`; they do not expose `referencePath`. The
Agent selects zero or one candidate, reads its guide, views its examples, and
derives a deck-specific Design Grammar before composing new pages.

### Preserve the eight IDs but recreate the assets

The seven bundled PPTX IDs and Grid ID remain searchable, minimizing needless
name churn in a 1.1 release. Their instructions and calibration images are new
OfficeKit-authored work with unrelated content and geometry. Grid becomes a
style guide about alignment and systematic rhythm; its 26 layouts and all
runtime modules are deleted.

### Route references outside the template concept

A user-provided PPTX is a reference deck or source-continuation input. A brand
document is a design-system authority. Neither is cataloged as a template.
When a design system conflicts with a selected template, the design system
wins and the template is not blended. No candidate is a valid successful
selection.

## Risks / Trade-offs

- [Style examples are too weak to guide a fresh Agent] → Require four to six
  examples spanning at least three declared page roles and dogfood one unrelated
  deck before accepting the creator workflow.
- [Schema v2 presentation Skills survive in user directories] → Reject them
  with a direct specialist-creator migration message; do not silently skip or
  materialize them.
- [Clean-room recreation still looks derivative] → Use references only for
  abstract traits, replace names/content/geometry, and review new examples
  without packaging old screenshots or files.
- [Mixed v2/v3 catalog logic becomes another dual template system] → The split
  is only by Office artifact kind: PowerPoint has one v3 form; DOCX/XLSX retain
  their existing form outside this change.

## Migration Plan

1. Land the schema and specialist Creator behind the current package source.
2. Update search and Presentations routing, then reject presentation schema v2.
3. Rebuild and review each of the eight style Skills independently.
4. Remove all old PPTX, Grid code, previews, fallback guidance, and package
   evidence in one cleanup milestone.
5. Release `1.1.0` after one final repository/package verification pass.

Rollback is a normal revert before release. After release there is no runtime
compatibility rollback path because the removed representation is deliberately
unsupported.

## Open Questions

None. The single format, eight-style rebuild, lack of compatibility layer, and
1.1.0 release choice are explicit product decisions.
