# Design: bounded 48-point compact projection

## Boundary

`PpjLinePathCodec.TryProjectKimi` will accept a literal single stroked path
whose cubic chain is exactly the deterministic equal-parameter lowering of a
Kimi high-degree `smooth` point list. The existing Householder QR solve,
integer quantization tolerance, and formula proof remain unchanged.

Only `MaxSmoothProjectionPoints` changes from 32 to 48. This adds support for
33 through 48 point lists while retaining a finite matrix size. Paths above
48 points, unstable reconstructions, and arbitrary cubic chains remain typed
paths.

## Evidence

The regression will author a 33-point `line` with `curve: "smooth"`, compile
it, remove the embedded PPJ snapshot, project the resulting PPTX, and assert
that the result still contains compact `points` and `smooth` rather than a
typed `path`.

## Non-goals

- Raising the authored 128-point input budget.
- Accepting arbitrary multi-segment cubic paths as compact points.
- Adding guide, handle, formula, multi-path, closed-fill, or group-transform
  geometry semantics.
- Replacing the bounded numeric proof with lossy approximation.
