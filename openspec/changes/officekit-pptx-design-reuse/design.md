## Context

An imported PPTX has two different kinds of information: exact package bytes
that must survive unrelated edits, and semantic clues that help an Agent decide
how to continue the deck. Treating those clues as a new universal object model
would recreate the fidelity problem. The profile therefore reports evidence,
while the existing Presentation projection and capability-issued Edit Plan
remain the only mutation paths.

## Decisions

### 1. Source evidence first

The profile binds to the source SHA-256 and records structural part hashes for
the presentation, masters, layouts, and themes. It counts package features and
inspected semantic elements without embedding source bytes. A profile generated
for another revision is stale and cannot authorize an edit.

### 2. Separate design language from edit capability

Palette and typeface frequencies, normalized geometry rhythm, slide archetypes,
and repeated component signatures are descriptive evidence. They are not
permissions. `Presentation.inspect({ includeComponentCandidates: true })` can
issue defensive `componentCandidate` records with a source revision,
occurrence IDs, and ownership evidence. V1 is deliberately inspect-only:
ambiguous, opaque, and relationship-bound graphs are marked blocked instead of
being lowered into an unsafe partial component operation.

### 3. Keep opaque content visible

Native objects are summarized by kind, slide, and stable IDs. OLE, SmartArt,
WPS extensions, animations, and other unknown graphs remain opaque unless a
codec capability specifically proves an operation. The profile never flattens
them into shapes or images to make generation easier.

### 4. Start with evidence, then add reuse

The first implementation is the deterministic `pptx-design-profile` evaluator
and three-sample evidence file. The bounded source-derived slide reuse slice
now copies complete ownership-checked graphs, appends the copy, reopens it, and
permits a typed continuation edit before a second package review. The same
three samples are rendered through LibreOffice and Poppler; every original
page is pixel-identical after the appended continuation page, and the new page
is required to render non-blank. Source-bound review can compare an edited
artifact with its pre-edit baseline and downgrade only exact pre-existing
semantic/layout issues; structural package failures and new errors remain
blocking. A deterministic
public-REPL rehearsal now runs three fresh task workspaces through two commits,
two resumes, verification, and publish without changing the source. It is
evidence for the task protocol, not a model black-box score. The change still
keeps component mutation and native-host acceptance as later gates rather than
turning visual similarity into an unproven mutation authority.

## Non-Goals

- No universal PPTX JSON/AST or conversion to HTML/PPTD.
- No semantic inference that depends on a model, image-generation tool, or
  visual host.
- No automatic cloning based only on visual similarity.
- No Windows acceptance claim from LibreOffice or a static profile.
