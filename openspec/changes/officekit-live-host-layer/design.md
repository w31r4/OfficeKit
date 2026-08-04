# Design

The bridge is one local HTTPS/session implementation with an adapter selected
when the bridge is started. The adapter owns static assets, client descriptor
validation, operation validation, result envelopes, labels, and audit
summaries. Every session is bound to one host and one add-in pane; the bridge
does not let a PowerPoint request reach Excel or vice versa.

The PowerPoint add-in uses Shared Runtime with a long-lived task pane. It
polls the bridge, executes a fixed operation set through PowerPoint Office.js,
and returns bounded results. Object IDs and optional expected text snapshots
make stale edits explicit. The add-in has no arbitrary JavaScript execution
operation.

The CLI and REPL are control-plane and task-plane facades respectively. Both
load the PowerPoint client lazily, so root imports, REPL startup, template
search, and ordinary file tasks do not create certificates, start a bridge, or
load Office.js.

Windows x64 manual evidence is required before coverage can call PowerPoint
Live `done`; macOS automated checks are not substituted for that evidence.
