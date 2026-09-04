# Author report

Case: `academic-research-10`  
Fixture: `nasa-mms`  
Delivery: reader  
Status: `passed-with-limitations`; no source-bound edit was authorized.

## Task-local style brief

Goal: let a research reader inspect the training method, its provenance, and the telemetry-window evidence without overstating a result.

Visual thesis: preserve the source page's figure-first methods explanation and keep the source attribution adjacent to the method claim.

Primary carrier: mixed source-native diagram, explanatory text, and source telemetry image.

Visual move: retain the source's left-to-right method flow and the lower telemetry-window comparison; no new geometry was added.

Reading order: 1) title and source attribution, 2) vector-distance/TCN method diagram, 3) context/suspect-window labels and explanatory copy; the footer remains secondary.

Visual DNA: source template styling—dark space header, white evidence field, light-blue footer, precise black linework, and restrained blue plot accents.

Avoid: invented statistics, unsupported result language, generic overlays, source flattening, and edits to other pages.

Review probes: requested semantic surfaces present; chart labels/axis content legible at final size; no evidence-bearing object obscures another.

## Source and target

- Source path: `/Users/zfang/workspace/open-office-artifact-tool/tmp/reference-pptx-downloads/nasa-mms-machine-learning.pptx`
- Source SHA-256: `531c82797fde09b1ebe1e868ca9cd44c3e2f675dc8f09f58b54bab6a62629723`
- Existing PPJ: `outputs/deck.ppj`
- Assigned page: `page-ppt-slides-slide9.xml` (presentation page 9; zero-based page index 8)
- Source revision: `pptx-531c82797fde09b1`
- Source-bound projection: 21 pages, 329 expanded elements, 53 assets

## Serial stage decisions

### Stage 1 — native note/result-label edit

Fail closed. The assigned page contains no sample-size note and no result label. Its visible text is the methods title, Carmona et al. attribution, vector labels, window labels, and explanatory copy. The case supplies no replacement sample-size or result-label values. No text leaf, native reference, or other page was changed.

### Stage 2 — table/chart alignment and label repair

Fail closed. The assigned page has no typed `table` or `chart` element. It has one source-bound raster telemetry plot (`page-ppt-slides-slide9.xml-element-2`) plus source-bound labels and shapes; the requested interval/footnote/table surface is not present. The rendered chart-like image and its context/suspect labels were legible enough at delivery size, so no frame, crop, label, overlay, or z-order edit was made.

No new shape, overlay, data value, conclusion, citation, page, or source part was invented or moved.

## Stable IDs preserved

Target page ID: `page-ppt-slides-slide9.xml`.

Target-page element IDs, retained in order:

`page-ppt-slides-slide9.xml-element-1`, `-2`, `-3`, `-4`, `-5`, `-6`, `-7`, `-8`, `-9`, `-10`, `-11`, `-12`, `-13`, `-14`, `-15`, `-16`, `-17`, `-18`, `-19`, `-20`, `-21`, `-22`.

The reimport audit found no differences in page IDs, target element IDs, or target native object hashes. Non-target pages and parts were not edited.

## Command log and hard gates

Commands were run with the public OfficeKit CLI at `node /Users/zfang/workspace/officekit-main-skill-eval-20260903/bin/officekit.mjs ppj`.

1. `inspect outputs/deck.ppj --page page-ppt-slides-slide9.xml --json` — confirmed the target page and issued capabilities.
2. `render outputs/deck.ppj -o outputs/previews-baseline --pages 8-10 --json` — baseline visual inspection.
3. `check outputs/deck.ppj --json` — passed; `changedNodeIds: []`; expected opaque-content warning only.
4. `build outputs/deck.ppj -o outputs/deck.pptx --json` — passed; `changedParts: []`; `changedNodeIds: []`.
5. `render outputs/deck.ppj -o outputs/previews --pages 8-10 --json` — passed; rendered pages 8, 9, and 10.
6. Manual inspection of `outputs/previews/slide-008.png`, `slide-009.png`, and `slide-010.png` — no target-surface collision or repairable occlusion found; no repair loop was needed.
7. `review outputs/deck.ppj --json` — `passed-with-limitations`; structural playback evidence only; automated visual review unavailable.
8. `import outputs/deck.pptx -o evidence/reimport.ppj --json` — passed; source hash, page count, and expanded element count preserved.
9. Reimport inspection and ID/native-object-hash diff — no differences reported.

## Hashes and evidence

| Artifact | SHA-256 |
| --- | --- |
| `outputs/deck.ppj` | `824421f73c1f979d6dddc099e6e0ecfd9ac7255b49d3c17eaa5d1830162ee0fd` |
| `outputs/deck.pptx` | `531c82797fde09b1ebe1e868ca9cd44c3e2f675dc8f09f58b54bab6a62629723` |
| `evidence/reimport.ppj` | `0aed428b6573d969f0a6eb7d07649f8816c92cc6b1cea6e32ff593edba93b8a2` |
| `outputs/previews/slide-008.png` | `2e51614976a41e664504a2b7741fe88a6d294742cf58ab886952a4c41b269a91` |
| `outputs/previews/slide-009.png` | `5a756915693b562a79cc0d39338d8cfbe2570c0c4a0ed8a6c79aeada47755813` |
| `outputs/previews/slide-010.png` | `c498ce336efaed7cd65ae3aa2dec49021fb018a3dfedeb8eb87f80cf094da9de` |

The built PPTX is byte-identical to the supplied source hash, consistent with a no-op source-bound build. Review reported the pre-existing limitations: 262 retained opaque OPC parts, several pre-existing out-of-canvas frames, existing motion-density warnings, no automated visual review, structural rather than desktop playback evidence, and no publication path supplied to the review command.

## Limitation

The requested content was intentionally not moved to another page. Both stages remain unchanged because the assigned page lacks the named editable surfaces and the case supplies no missing values. The deliverable therefore proves source preservation and round-trip integrity, not completion of the unavailable semantic or table/chart edits.
