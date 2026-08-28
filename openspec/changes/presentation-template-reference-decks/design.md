# Design

## Published surface

```text
artifact-template-<id>/
├── SKILL.md
├── artifact-template.json   # schemaVersion 4
├── agents/agent.yaml
└── assets/
    ├── reference.pptx       # OfficeKit-authored, hash-bound
    ├── preview.png
    └── examples/*.png
```

The guide remains the style authority. The reference deck is a native,
inspectable calibration/source asset: an Agent may inspect its layer order and
selectively reuse verified evidence, but must compose new pages rather than
clone coordinates. Source decks, extracted media, and QA evidence remain in
the task.

## Creator contract

The creator requires an absolute `referencePath` ending in `.pptx`, checks the
bounded file as a structurally valid Open XML package, writes it atomically,
and records `provenance.referenceSha256`. Updates continue to require the
current sidecar hash. The creator does not infer or download a reference deck.

## Search contract

Schema v4 is valid only for presentation templates and requires the reference
asset plus its hash. Search returns `referencePath` for v4 candidates while
retaining read-only support for existing v3 entries during the migration
window. A v3 entry is legacy evidence and cannot be used to claim a completed
95/95 restoration.

## Restoration indices

Migration evidence records two independent 0–100 scores:

- visual: silhouette, hierarchy, palette/surfaces, typography, density,
  visual carriers, layer relationships, motifs, and example coverage;
- functional: inspect discovery, editable leaves, reusable assets, round-trip
  stability, native rendering, background/layer fidelity, opaque preservation,
  and safe refusal.

Scores must cite the source image or render, inspect/edit/re-import output, and
the relevant hashes. Both scores must be at least 95 before a template is
labelled restored. The repository-only
`scripts/presentation-template-fidelity.mjs` command applies these fixed
weights and requires source, render, inspect, edit, re-import, and package
evidence before reporting `restored: true`. Missing evidence is not a zeroed
aesthetic judgement; it is an incomplete migration.
