## 1. Strike field and propagation

- [x] 1.1 Add the chart strike schema/wire/native style field, preserving explicit noStrike and rejecting invalid values/global-font-only misuse. Verify focused native style and protobuf presence assertions.
- [x] 1.2 Propagate strike through authored/source-bound mapping, projection, rich styles, field precedence and vector text. Verify line/combo creation/change/cancellation/deletion/recreation, exact no-op/non-target ZIP preservation and one vector override fixture.

## 2. Regression and discoverability

- [x] 2.1 Keep unsupported native properties opaque and preserve wire strike through unrelated JS edits. Run focused native/JS regressions and proto:check on the isolated snapshot.
- [x] 2.2 Update chart guidance, registry, coverage and F-07 backlog; regenerate the PPJ manual/matrix and pass affected drift, portability, strict OpenSpec and diff checks. Record host/runtime evidence limits.

Evidence (2026-09-09): isolated snapshot based on ade46bd5 passed 23/23 focused native regressions with SDK 8.0.128 (PpjChartTextStrike, PpjChartTextLanguage, PpjTrendlineRichText, PpjTrendlineLabel, PpjChartTextAlignment/Underline/Fill). Four strike cases prove line/combo lifecycle, native attribute cancellation versus absence, non-target ZIP preservation, invalid values/global-font restriction, field precedence and vector title/label output. Unsupported native strike/unknown character attributes keep analytics closed. JS worksheet chart wire preservation, proto:check against regenerated/staged binding, strict OpenSpec validation, maintainer, matrix drift, preview capability coverage and Skill portability (255 files) passed in this worktree; reference sync (333 files) passed in the main worktree with its local legacy inputs. No NativeAOT release build or PowerPoint host/visual acceptance was performed. Full F-07 and the overall backlog remain open.
