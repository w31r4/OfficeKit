# Default Template Library

The catalog ships 22 independently selectable Office templates: 7 document
templates, 9 presentation visual grammars, and 6 spreadsheet templates.

Presentation entries use one clean-room format:

```text
artifact-template-<name>/
├── SKILL.md                 # the style grammar and authoring route
├── artifact-template.json   # English retrieval metadata and hashes
├── agents/agent.yaml
└── assets/
    ├── preview.png
    └── examples/slide-01..05.png
```

A presentation template is deliberately not a source deck. It contains a
direction, composition rules, layering guidance, and representative original
images. It contains no PPTX, MJS, DSL, fixed layout, or cloneable page. Read the
guide, inspect the examples, then derive a deck-specific design grammar and
compose editable pages with the Presentations Skill. Examples communicate range;
they are not material to copy.

The remaining document and spreadsheet entries retain their MIT-licensed
reference files and their exact source hashes. Every shipped binary has an
integrity record in `integrity.json`; see [LICENSE.md](LICENSE.md).

Search the catalog with `officekit template search`. It returns candidates
only; the Agent chooses zero or one presentation grammar. User-provided
templates and design systems take precedence. No template is copied into a
project by `officekit init`.

For an existing presentation, use the source-bound import/edit route. A
clean-room presentation template is for new composition, not for pretending
that an old reference file is still present.
