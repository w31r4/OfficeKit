# Evidence

## Implementation

- PPJ accepts `bubbleSizeScale: sqrt | linear | log` and a strictly increasing
  `bubbleRadiusRange` from 2 through 72 points.
- The NativeAOT authored compiler uses one shared positive size domain across
  all bubble series and emits stable editable ellipse, axis, title and legend
  objects.
- Ordinary bubble charts without either field remain native ChartParts. The
  explicit mapping fails closed for native chart-build animation, formulas,
  secondary axes, per-series ranges and unsupported source-bound mutation.

## Lean verification

- NativeAOT codec build passed with zero warnings and zero errors after the
  compiler and semantic-validator implementation.
- One exploratory execution of the existing comprehensive authored PPJ
  contract accepted the new semantic validation and completed PPJ compilation;
  it then stopped at the deliberately changed expanded-element count. The
  temporary assertions were removed instead of making that oversized test
  heavier or rerunning its eleven-minute path.
- Presentation Skill maintenance check passed with 151 Help APIs, 80 native
  leaves and 13 host-only operations after regenerating `ppj.md`.
- The OpenSpec change passed strict validation before implementation. The
  isolated worktree has no local OpenSpec executable, so the same validation
  was not repeated after prose-only task updates.
- `git diff --check` passed.

## Honest boundary

This slice proves the language contract, semantic boundary, compiling code and
Agent-facing discoverability. It does not claim a completed end-to-end visual
round trip, Windows PowerPoint playback, imported native ChartPart mutation or
reverse inference from an arbitrary DrawingML group. Those claims require a
later real deck, not more synthetic test infrastructure in this slice.
