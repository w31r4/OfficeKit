# Add paragraph default reflection rotate-with-shape leaf

The paragraph `defaultText` reflection profile already reads the native
`rotWithShape` flag, but it cannot expose or edit that one source-owned value.
Add the missing `textDefaultReflectionRotateWithShape` native leaf at
`paragraph.style.defaultText.reflection.rotateWithShape`. Keep the experiment
source-bound and token-only: an explicit native `0` or `1` is required, only the
owning SlidePart changes, and unsupported effect graphs remain opaque or fail
closed.

This closes one P0 text-effect leaf without changing the wire protocol or
claiming PowerPoint host rendering equivalence.
