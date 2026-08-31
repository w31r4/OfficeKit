---
name: office-kit
description: Use this skill to plan and coordinate broad, ambiguous, cross-format, or multi-deliverable Office work across Word/DOCX, Excel/XLSX/CSV, PowerPoint/PPTX, and PDF, including deciding whether to use zero or one available template. Load only the required installed domain Skills and preserve their own edit and QA rules. Do not use when the user explicitly invokes a single domain or template Skill for a self-contained task.
---

# OfficeKit

Turn one Office request into a small, explicit artifact workflow. Route each
output to its owning Skill, load only the instructions needed for that route,
and preserve the owning Skill's safety and QA rules.

Use the public `office-kit` package for OfficeKit work. Never import or use
`@oai/artifact-tool`: it is a different host-bundled runtime, not an OfficeKit
alias or fallback, and its output must never be attributed to OfficeKit.

When maintaining a Presentation capability rather than producing an artifact,
load `presentation-skill-maintainer`. It keeps PPJ Schema, native compiler
ownership, Help, Agent guidance, review rules, and examples synchronized.
Other formats continue to follow their owning Skill and repository guide.

For a multi-step task, use the portable [task REPL workflow](references/repl.md).
Inspect `officekit tasks --json` before continuing prior work; open the matched
task or create one explicitly. Import only the selected domain API with
`ctx.import`, keep reusable helpers or live objects in `ctx.state`, and treat
them as process-local. Stage inputs with `ctx.input`, commit only a candidate
whose post-edit review passes, and publish only from the current reviewed
commit. Setup commands
such as template search, provider installation, and Excel add-in installation
remain explicit; they are not hidden inside a cell.

Treat `taskRoot` as a private store. Do not use shell or Node filesystem tools
(`cat`, `sed`, `ls`, `file`, `find`, `rg`, or direct reads) on task manifests,
plans, candidates, revisions, sessions, or evidence paths. Use `officekit tasks`,
REPL responses, and public package APIs; a returned absolute evidence path is a
handoff location, not an instruction to inspect its private bytes.

Before planning paths or a visual asset, read [the workspace and evidence
contract](references/workspace.md) and [the visual capability matrix](references/capabilities.md).
Before delivering an edited artifact, follow the shared [post-edit review
contract](references/review.md).
Use `workspaceRoot`, `taskRoot`, `inputRoot`, `assetRoot`, `outputRoot`,
`evidenceRoot`, and a local `sessionId`; do not depend on a host chat, thread,
plugin, or image tool.

After an artifact is exported, use the [post-edit review contract](references/review.md)
to reopen it and separate semantic, structural, layout, optional text-reading,
visual/human, and delivery evidence. AnyDoc is the parser behind that optional
view, never a replacement for native OfficeKit inspection or pixel review.

## Respect explicit choices

- If the user explicitly invokes a domain, business, or template Skill for a
  self-contained task, use it directly.
- If an existing Office file is itself the output being edited, treat that file
  as the reference for that output. Do not search for a decorative template for
  it; decide independently for any other new outputs.
- Do not replace an unavailable Skill, template, provider, or live application
  with a different execution path. Report the missing component.

## Build the artifact route

1. Inventory every input file and requested output.
2. Identify whether the task is read-only, creates a new artifact, edits an
   existing artifact, or converts between formats.
3. Assign exactly one owning Skill to each output:
   - DOCX or Google Docs handoff: Documents
   - XLSX, CSV, TSV, or Google Sheets handoff: Spreadsheets
   - an already-open Excel workbook: Excel Live Control
   - PPTX or Google Slides handoff: Presentations
   - an already-open PowerPoint presentation: PowerPoint Live Control
   - PDF: PDF
4. For multiple outputs, order the owners as a dependency graph. Pass facts,
   tables, images, and structured content between steps; never let two Skills
   mutate the same file.
5. Read [routing.md](references/routing.md) for cross-format work, live Excel,
   live PowerPoint,
   missing Skills, or an ambiguous owner.

## Load only the selected Skill

Load the installed `SKILL.md` for each chosen owner before using its package
APIs or scripts. Follow that Skill's import, edit, source-preservation, render,
verify, and publication rules without weakening them.

When OfficeKit is active, carry the routed task through the selected owners.
Do not send the user away to repeat the request through another Skill.

Do not preload every Office Skill. Do not copy domain API documentation into
this coordination layer.

## Keep net-new PPT work conversational

For a net-new PPTX or a broad deck redesign, load the Presentations Skill and
author one `.ppj` program. Keep its communication brief, narrative, per-deck
design grammar, page jobs, and editorial rules in that program.
Return the first checked deck as a guided working draft, revise the latest
inspected draft in natural language, and publish only after acceptance unless
the user requested one-pass final delivery. Read-only and narrow edits stay
  direct. Discuss the draft, changes, and final deck, not Skills, routing, CLI/parser mechanics, or QA internals unless asked or blocked.

## Decide whether a template helps

Make a template decision only for a new or substantially redesigned DOCX,
XLSX, or PPTX.

1. Decide whether the artifact goal is clear and whether a template is already
   specified. Search only when the goal is clear and no template is specified;
   otherwise clarify the goal, inspect the specified template, or proceed
   directly as described in
   [template-selection.md](references/template-selection.md).
2. Treat a user-provided template as a task-scoped reference, not a catalog
   entry. Never register it unless the user explicitly asks for later reuse.
3. Summarize the purpose, audience, content shape, visual traits, and required
   operations as short English search terms. Continue speaking to the user in
   the user's language.
4. Run `officekit template search ... --json`; do not inspect every template
   file.
5. Choose exactly one of `selected`, `ask`, or `none`.
6. Load previews only for the final one to three candidates.
7. Before selecting a DOCX/XLSX template, load the owning domain Skill and
   confirm that requested edits fit its verified edit profile. For PPT, load
   the selected Template Skill and representative examples, derive a new
   deck-specific Design Grammar, and author the current content in PPJ. If the
   candidate declares a clean-room reference PPJ/PPTX, inspect or import it
   through the ordinary PPJ/source-bound route; it is reusable evidence, not a
   fixed page recipe or an edit profile for an external source deck.

`none` means the owning Skill should compose the artifact from first
principles. It is a successful design decision, not an error or fallback.

Do not use the Office template catalog for a PDF-only task. When PDF is the
final form of an Office artifact, apply the template to the Office source step,
then let the PDF Skill inspect and verify the final PDF.

## Execute and verify

- Protect every input and retained template from overwrite.
- For durable presentation work, follow `tasks → repl → plan → input → edit →
  review → commit → publish`; other artifact tasks may omit `plan`.
  Every `ctx.commit(candidate, options)` requires a concise non-empty
  `options.summary`. A failed review remains attention and cannot replace task
  HEAD. Put a multi-line Agent-authored cell in a regular `.mjs` file and run
  it with `officekit repl ... --file <cell.mjs>`; use raw JSONL only when one
  live process genuinely needs several cells.
- Complete each artifact under its owner's workflow.
- For a conversational PPT draft, run the draft checks defined by the
  Presentations Skill and return the working path plus draft guide. Do not call
  it final or publish it as the delivery artifact.
- Reopen or reimport the result when the owner requires it.
- Run the owner's semantic, structural, layout/render, and delivery checks in
  the order defined by the post-edit review contract.
- Request the lazy text reading view (`contentView: "anydoc"`) only when it can
  close an identified text or table content-coverage gap. It is not a
  visual-review result and does not resolve OCR, layout, image, formula, or
  metadata-provenance gaps.
- For a multi-artifact task, verify shared facts, numbers, names, dates, and
  visual identity across outputs.
- After finalization, return the final files and any explicit capability
  limits. Keep the internal route and template decision out of the normal
  user-facing response unless the user asks for them.

Return each artifact as an absolute path with its type and SHA-256. Include
page, slide, sheet, or range locators and inspect/render/verify evidence paths
when available. Report `visualReview: "complete"`, `"unavailable"`, or
`"requires-human"` according to the capability matrix. The active host can
turn these paths into links; this Skill does not emit host-specific citation
syntax.

Do not describe an unverified output as complete.
