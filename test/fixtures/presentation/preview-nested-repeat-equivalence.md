# Nested repeat / slot preview pair

`preview-nested-repeat-equivalence.json` is consumed by
`test/ppj-preview-scene-native.mjs`. Both programs start from the checked-in
`examples/ppj/minimum.ppj`; the test does not derive the explicit answer from
compiler output.

The high-level program contains an outer component repeated horizontally twice.
Each instance contains an inner component with an orange rectangle and a supplied
text slot containing the literal `0`. The explicit program independently specifies
four elements at the expected positions. Both include the same line chart with
observations `[1, null, 0]`.

Required evidence:

- Four shape/text frames match the explicit coordinates; both orange interior
  pixels and a blank point between repeats match their expected colors.
- The chart keeps missing index 1 distinct from true zero, preserves both isolated
  observations, and does not invent a connecting segment.
- Ordered, typed native content payloads match byte for byte. The outer element
  wrappers' generated IDs and provenance are deliberately not compared as visual
  payload; the test separately checks source paths, generated attribution,
  distinct instance identities and distinct scene addresses.
- Full raw page rasters match, including text and chart output, not merely the two
  sampled orange pixels. This is same-backend equality, not host-PowerPoint QA.
- Input programs remain unchanged. Fixture bytes are included in the integration
  report's input/implementation identity check before and after execution.
- Scene-on and ordinary scene-free compilation produce identical candidate bytes
  for each side; collecting component origins cannot change the exported file.

Compilation or comparison failures remain fatal after independent regressions
run. A missing or failed side cannot pass equivalence.

The suite retains the plain pair and also runs a styled pair, composing this
geometry fixture with `preview-style-grammar-equivalence.json`. The inner tile
uses the named shape style; the definition-supplied slot uses the named text
style with an explicit false override. The explicit side uses independently
specified resolved styles. Both repeated slots must carry DejaVu Sans, 32pt,
non-bold text and actual blue glyph pixels, in addition to all origin, instance,
geometry, missing-data and full-raster comparisons above. Both fixture digests
are recorded. The small slot's glyph extent is inspected beyond its frame;
this proves equality, not that text overflow/layout is fully implemented.

These pairs cover one nested-repeat/slot arrangement. Other repeat layouts,
full inherited styles and source-bound edit equivalence require their own cases.
