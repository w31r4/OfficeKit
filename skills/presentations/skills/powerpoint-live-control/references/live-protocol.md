# PowerPoint Live protocol 1

Requests are JSON objects with these required fields:

```json
{
  "protocol": 1,
  "sessionId": "powerpoint-…",
  "idempotencyKey": "one-operation-key",
  "operation": "read_slide",
  "args": { "slideId": "slide-1" }
}
```

The typed operation set is `read_presentation`, `read_slides`, `read_slide`,
`read_selection`, `write_text`, `add_textbox`, `add_shape`, `add_image`,
`update_shape`, `delete_shape`, `add_slide`, `read_slide_image`, and `save`.
The bridge bounds requests at 10 MB, responses at 10 MB, images at 8 MB, and
geometry at 100,000 points. Data URLs must be bounded PNG, JPEG, GIF, or safe
SVG images. `write_text`, `update_shape`, and `delete_shape` may carry an
`expectedSnapshot` containing the target ID, name, type, text, or geometry;
any mismatch returns `stale-target` before the mutation.

Success returns `{ protocol: 1, ok: true, result, audit }`. Failure returns
`{ protocol: 1, ok: false, error: { code, message, retryable, maybeApplied } }`.
After a timeout or disconnect, `maybeApplied` requires rereading the target
before any retry. `unsupported-capability`, `stale-target`, and
`session-unavailable` are explicit outcomes; they do not invoke a closed-file
PPTX edit or arbitrary Office.js code.
