# Presentation template archive move

Date: 2026-08-31

## Purpose

Separate three things that had been easy to confuse:

1. current PPJ authoring and active Template Skills;
2. visual and editable template examples used for research;
3. retired MJS authoring code and Grid-era workflows.

This is a repository reorganization. It does not change the Presentation file
model, wire protocol, template search results, or generated PPTX behavior.

## Path map

| Previous location or source | New location | Status |
|---|---|---|
| `reference/office-artifact-tool` | `legacy/presentations/mjs/office-artifact-tool` | pinned legacy/reference submodule |
| user-provided `设计系统模板库-30套风格` | `examples/presentation-template-references/kimi` | 30 unpacked design/image reference pairs |
| seven historical Codex PPT template packages | `examples/presentation-template-references/codex` | 7 unpacked guides, metadata, previews, and PPTX files |
| old Presentation MJS/Grid implementation inside the pinned submodule | `legacy/presentations/mjs/office-artifact-tool/skills/presentations` | legacy only |

The submodule remains pinned at
`73c99c67ca7bbaa82cec0b158c647db583dcd970`. Reference-sync tooling now reads
the same immutable source from its legacy path; the content and recorded
digests did not change.

## Active and inactive boundaries

The following remain active and were not moved:

- `src/presentation/*.mjs`: current JavaScript runtime and compatibility code;
- `skills/presentations`: current PPJ-facing Presentation Skill;
- `skills/presentation-template-library`: current Presentation Template Skills;
- DOCX and XLSX templates under `skills/default-template-library`.

The following are examples only and must not enter runtime discovery or package
selection:

- `examples/presentation-template-references/kimi`;
- `examples/presentation-template-references/codex`.

They are also explicitly excluded from the npm package. The ordinary runnable
examples remain published.

The following are legacy only and must not be extended as product code:

- `legacy/presentations/mjs`;
- the Grid MJS modules and former Presentation Skill contained there.

## Content inventory

- Kimi references: 30 `design.md` files and 30 `reference.jpg` files.
- Codex references: 7 `reference.pptx` files, 7 previews, and their historical
  Skill/metadata files.
- Kimi application binaries, PPTD runtimes, private payloads, generated MJS,
  caches, and temporary outputs: not copied.

## Future maintenance

New PPJ primitives and authoring guidance belong in the current schema,
compiler, Help, and Presentation Skill. New template research material belongs
under `examples/` until it has a clean-room guide, original calibration deck,
rights record, and review evidence. Nothing should be promoted from `legacy/`
by path alone.
