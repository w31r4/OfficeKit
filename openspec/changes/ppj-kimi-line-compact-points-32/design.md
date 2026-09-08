# Design: bounded 32-point compact projection

## Boundary

`PpjLinePathCodec.TryProjectKimi` will accept a literal single stroked path
whose cubic chain is exactly the deterministic equal-parameter lowering of a
Kimi high-degree `smooth` point list. The interpolation observations, integer
quantization tolerance, and formula proof remain unchanged. The bounded
control solve uses Householder QR rather than normal equations so extending
the degree does not square the interpolation matrix's condition number.

Only `MaxSmoothProjectionPoints` changes from 24 to 32. This adds support for
25 through 32 point lists while retaining a finite matrix size and the
existing fail-closed behavior for unstable or arbitrary paths.

## Evidence

The regression will author a 25-point `line` with `curve: "smooth"`, compile
it, remove the embedded PPJ snapshot, project the resulting PPTX, and assert
that the result contains the original compact `points` and `smooth` marker
without a typed `path`. A second projection check is unnecessary for a
non-editing import proof; the existing six-point source-bound and authored
round-trip tests continue to cover edit and lower-degree behavior.

## Non-goals

- Raising the authored 128-point input budget.
- Accepting arbitrary multi-segment cubic paths as compact points.
- Adding guide, handle, formula, multi-path, closed-fill, or group-transform
  geometry semantics.
- Replacing the bounded numeric proof with lossy approximation.
