---
name: "powerpoint-live-control"
description: "Control a PowerPoint presentation already open in desktop PowerPoint through OfficeKit's local typed Add-in. Use for the current open deck, selection, or unsaved presentation changes. Use Presentations for standalone PPTX files."
---

# PowerPoint Live Control

Use this Skill only when the user explicitly targets a presentation already
open in desktop PowerPoint. A standalone `.pptx` belongs to the Presentations
Skill and must not be silently switched to the live route.

## Connection and work loop

Use the typed `officekit live` commands. Select exactly one intended session,
then execute one validated operation at a time. Re-read the target after every
mutation. If a result contains
`maybeApplied: true`, inspect the slide before retrying. Disconnect when the
task is complete.

The one-time setup is explicit:

```sh
officekit live install --app powerpoint --yes --json
officekit live doctor --app powerpoint --json
officekit live sessions --app powerpoint --json
officekit live execute request.json --json
officekit live disconnect SESSION_ID --json
```

Upload the printed manifest in PowerPoint through **Home > Add-ins > My
Add-ins > Upload My Add-in**, then open OfficeKit from the Home ribbon and
connect the intended presentation.

## Typed operations

The protocol supports presentation and slide summaries, current selection,
text replacement, text boxes, basic shapes, images, geometry updates, shape
deletion, new slides, bounded slide PNG previews, and explicit `save`.
Requests use stable `slideId`/`shapeId` locators and may include expected text
to prevent stale edits. There is no arbitrary Office.js execution entry point.
See [references/live-protocol.md](references/live-protocol.md) for the bounded
envelope and error fields.

Master/Layout graphs, SmartArt, animations, transitions, complex comments,
OLE, unsupported chart graphs, and other unverified host objects return
`unsupported-capability` or remain read-only. The Skill never falls back to a
closed-file edit while the user is targeting an unsaved live presentation.

## Review and save

After a visual change, request `read_slide_image`, register the returned image
under `evidenceRoot`, and report the visual review state. Saving is always an
explicit operation; OfficeKit does not choose a path or overwrite a file.

Return the session result, operation audit, evidence paths, and any
`unsupported-capability` or `maybeApplied` state. Keep presentation content out
of audit records.
