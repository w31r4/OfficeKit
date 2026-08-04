# PowerPoint Live

OfficeKit Live is a local bridge for an Office document that is already open
in the desktop application. Excel remains available through its historical
`officekit excel ...` commands; PowerPoint uses the common `officekit live ...`
surface:

```text
officekit live install --app powerpoint --yes --json
officekit live doctor --app powerpoint --json
officekit live sessions --app powerpoint --json
officekit live execute request.json --json
officekit live disconnect SESSION_ID --json
```

`install` creates a per-user certificate, bridge configuration, and the
PowerPoint add-in manifest. Upload that manifest once in desktop PowerPoint
with **Home > Add-ins > My Add-ins > Upload My Add-in**, then open OfficeKit
from the Home ribbon and connect the intended presentation. The bridge binds
only to loopback HTTPS, records content-free audit summaries, exits after its
idle grace period, and does not relay through a cloud service.

PowerPoint Live requests are protocol 1 JSON envelopes with a `sessionId`, an
`idempotencyKey`, a typed `operation`, and validated `args`. The V1 operations
cover presentation/slide summaries, current selection, text and basic shape or
image edits, slide creation, bounded slide PNG previews, and explicit save.
Mutations can include an `expectedSnapshot` for the target shape; a changed
text, geometry, type, or identity returns `stale-target` before the edit.
Each mutation should be followed by a reread; `maybeApplied: true` requires an
inspect-before-retry decision. Unsupported Master/Layout graphs, SmartArt,
animations, transitions, complex comments, OLE, and unverified chart graphs
return `unsupported-capability` or stay read-only. The live route never falls
back to editing a closed PPTX.

The `powerpoint-live-control` Skill is installed with `officekit init` and is
selected only when the user asks for the currently open presentation. Ordinary
PPTX creation and editing continue to use the Presentations Skill and the
OfficeKit/OpenChestnut file path.

## Acceptance boundary

The automatic mock and package gates run on every platform. The first real
host gate is Windows x64 desktop PowerPoint: manifest upload, pairing, two
isolated presentations, unsaved edits, selection, slide image review, explicit
save, reconnect, disconnect, and failure behavior. macOS is limited to build,
mock, and package checks until its separate desktop matrix is completed.
