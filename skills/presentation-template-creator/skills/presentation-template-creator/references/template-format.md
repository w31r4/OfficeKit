# PowerPoint template format

OfficeKit has one reusable PowerPoint template form. The published directory is:

```text
artifact-template-<slug>/
├── SKILL.md
├── artifact-template.json
├── agents/agent.yaml
└── assets/
    ├── preview.png
    └── examples/
        ├── 01-<role>.png
        └── ...
```

No other file is allowed.

## Guide body

The packaging script adds YAML frontmatter. Supply a Markdown body that explains
only this style's distinctive choices:

- communication territory and unsuitable uses;
- palette roles and surface hierarchy;
- typography roles and scale rhythm;
- geometry, lines, spacing, and density changes;
- image, chart, table, and diagram treatment;
- recurring motifs and variation limits;
- concrete anti-patterns;
- what each example demonstrates.

Do not repeat the general Presentations workflow or prescribe fixed coordinates.

## Packaging spec

```json
{
  "id": "artifact-template-example",
  "displayName": "Example",
  "description": "Create presentations with an evidence-led editorial style. Use when the user selects Example.",
  "guidePath": "/absolute/task/guide.md",
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

The resulting `artifact-template.json` is schema v3. Its hashes bind the guide,
preview, and every example. The generated preview is a deterministic two-column
montage; examples remain the full-resolution evidence the Agent should inspect.
