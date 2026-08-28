# PowerPoint template format

OfficeKit has one reusable PowerPoint template form. The published directory is:

```text
artifact-template-<slug>/
├── SKILL.md
├── artifact-template.json
├── agents/agent.yaml
└── assets/
    ├── reference.pptx
    ├── preview.png
    └── examples/
        ├── 01-<role>.png
        └── ...
```

No other file is allowed. `assets/reference.pptx` is an OfficeKit-authored,
hash-bound calibration/source deck; it is not the external reference that was
used to derive the style.

## Guide body

The packaging script adds YAML frontmatter. Supply a complete, self-contained
Markdown body that explains this style's distinctive choices. Do not shorten a
guide into a palette summary or an adjective list; the Agent must not have to
guess the design system:

- communication territory and unsuitable uses;
- palette roles and surface hierarchy;
- typography roles and scale rhythm;
- geometry, lines, spacing, and density changes;
- image, chart, table, and diagram treatment;
- recurring motifs and variation limits;
- concrete anti-patterns;
- what each example demonstrates;
- the preferred page archetypes and information density;
- how the style treats photographs, SVG, charts, tables, diagrams, and
  editable foreground/background layers;
- how an Agent may inspect or selectively reuse the packaged reference deck
  without treating it as a fixed layout or cloning every page;
- the final visual and structural checks that protect the style.

Do not repeat the general Presentations workflow or prescribe fixed coordinates.

## Packaging spec

```json
{
  "id": "artifact-template-example",
  "displayName": "Example",
  "description": "Create presentations with an evidence-led editorial style. Use when the user selects Example.",
  "guidePath": "/absolute/task/guide.md",
  "referencePath": "/absolute/task/reference.pptx",
  "useWhen": ["evidence led analysis"],
  "avoidWhen": ["playful consumer campaign"],
  "audiences": ["executive"],
  "contentShapes": ["findings", "evidence", "decisions"],
  "visualTraits": {
    "tone": ["editorial", "analytical"],
    "density": "medium",
    "colorMode": "neutral",
    "structure": ["narrative", "evidence led"]
  },
  "visualCommitment": "opinionated",
  "examples": [
    { "path": "/absolute/task/cover.png", "role": "cover" },
    { "path": "/absolute/task/evidence.png", "role": "analysis" },
    { "path": "/absolute/task/data.png", "role": "data" },
    { "path": "/absolute/task/decision.png", "role": "closing" }
  ],
  "provenance": {
    "license": "AGPL-3.0-or-later",
    "source": "OfficeKit original calibration work"
  }
}
```

Use English for search metadata. Example roles are `cover`, `section`,
`analysis`, `data`, `process`, `comparison`, `closing`, or `mixed`. Provide four
to six PNGs and at least three distinct roles.

The resulting `artifact-template.json` is schema v4 for presentation templates.
Its hashes bind the reference PPTX, guide, preview, and every example. The
generated preview is a deterministic two-column montage; examples remain the
full-resolution visual evidence the Agent should inspect. The reference deck is
the native, editable calibration source, while the guide remains the authority
for style decisions.
