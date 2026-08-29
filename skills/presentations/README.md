# Presentations

Presentations is the file-type wrapper plugin for presentation slide deck workflows.

This installable Skill bundle is distributed with `office-kit`.

## Included Skills

- `Presentations`: create, edit, render, verify, and export editable PowerPoint and Google Slides presentation decks.
- `Presentation Editorial Trim`: shape slide titles, visible support, labels, sources, and notes without changing facts or unrelated pages.
- `PowerPoint Live Control`: operate a presentation already open in desktop PowerPoint through the local typed OfficeKit bridge.

## Discoverability

Use this plugin for presentation-oriented terms from the file-type naming model: slides, deck, PowerPoint, Google Slides, presentation, presentations, PPT, and `.pptx`. Choose the Live Skill only when the user explicitly refers to the currently open desktop deck. Use Presentation Editorial Trim directly for copy-only work; the Presentations Skill also invokes it during creation and bounded edits.

The Presentation route is intentionally small. It sends authors to the PPJ
language manual and then to one focused design or source-bound reference. The
JSON Schema and native compiler own behavior; repository maintainers use the
bundled `presentation-skill-maintainer` to keep PPJ, Help, Agent guidance,
review rules, and examples synchronized.

## Source

The plugin tree is versioned directly under `skills/presentations` in the public repository.

## Compatibility status

Presentation guidance is source-bound and progressive: a style Template Skill,
a reference deck, and a source-continuation PPTX are separate authorities. The
API reference files document advanced package graphs; unsupported imported
topology remains opaque and fail-closed rather than being reconstructed.
