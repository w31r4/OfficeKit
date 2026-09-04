# Capability Coverage Contract

Scenario documents answer “what communication problem is this page solving?”
This ledger answers “which PPJ/native capability is exercised?” A case may
cover several rows; a row may be exercised by more than one case.

| Capability tag | Required evidence | Primary reference |
| --- | --- | --- |
| `text` | stable text ID, readable hierarchy, source/alt text when needed | `references/text.md` |
| `rich-text` | at least one run-level style or formula decision | `references/text.md`, `references/ppj.md` |
| `shape` | typed geometry, fill/stroke/opacity, no decorative default | `references/shapes.md` |
| `line-connector` | relation endpoints, arrow/line order, no evidence occlusion | `references/shapes.md` |
| `image-background` | image role, crop/focus, layer position, rights metadata | `references/image-sourcing.md`, `references/media-and-layers.md` |
| `mask-opacity` | deliberate mask/opacity and readable foreground | `references/media-and-layers.md` |
| `chart` | native chart or truthful vector chart, labels/axes and missing data | `references/charts-and-tables.md` |
| `table` | readable rows/columns, hierarchy without card wall | `references/charts-and-tables.md` |
| `group-z-order` | stable group identity and verified back-to-front order | `references/media-and-layers.md`, `references/imported-native-ref.md` |
| `formula` | bounded formula run when mathematical structure is evidence | `references/ppj.md` |
| `motion` | typed entrance/emphasis/transition with honest playback label | `references/motion.md` |
| `source-bound` | typed/nativeRef edit, target hash, non-target part preservation | `references/imported-native-ref.md` |
| `review` | check/build/render/review receipt and repairs after visible defects | `references/review-and-delivery.md` |

## Coverage rule

The frozen case manifest SHALL assign every required tag to at least one case.
If the current runtime cannot safely exercise a tag, the case is marked
`unavailable` before author runs; the experiment must not silently substitute a
different capability and call the row covered.
