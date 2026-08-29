## 1. Language foundation

- [x] 1.1 Add synchronized English and Chinese documentation explaining PPJ, Turing completeness, finite artifact state, browser control, and authored/imported authority.
- [x] 1.2 Add `office-kit/ppj/v1` JSON Schema with strict root metadata, intent, design, assets, source, components, pages, sections, custom shows, and comments.
- [x] 1.3 Define shared identity, frame, accessibility, style, color, asset, source, nativeRef, and opaque descriptor schemas.
- [x] 1.4 Define typed text, shape, image, chart, table, connector, group, media, placeholder, SmartArt, and OLE element schemas.
- [x] 1.5 Define bounded component parameters, slots, variants, repeats, conditions, stable expansion IDs, and hard budgets.
- [x] 1.6 Add one canonical authored PPJ fixture and path-specific schema validation coverage in the existing Presentation test surface.

## 2. Native PPJ model and validation

- [x] 2.1 Add additive proto v2 request/response messages and operations for PPTX projection and PPJ compilation.
- [x] 2.2 Implement C# PPJ root, deck, page, asset, source, component, and discriminated element models with strict JSON parsing.
- [x] 2.3 Implement semantic validation for IDs, references, paths, hashes, types, ordering, limits, and prohibited executable/native content.
- [x] 2.4 Implement deterministic component expansion with parameter, slot, variant, finite repeat, and simple condition support.
- [x] 2.5 Emit a normalized PPJ hash, expanded-node map, and path-specific diagnostics without mutating semantic input.

## 3. Authored PPJ compilation

- [x] 3.1 Map PPJ deck size, theme, design styles, sections, custom shows, comments, and page metadata into native Presentation state.
- [x] 3.2 Compile page backgrounds, ordered typed elements, groups, connectors, notes, transitions, and motion into existing native writer profiles.
- [x] 3.3 Compile rich text, local assets, crops, masks, charts, tables, data labels, and accessibility fields without a JavaScript Presentation model.
- [x] 3.4 Return a build receipt with PPJ hash, output hash, stable node mapping, asset mapping, expanded components, and package footprint.
- [x] 3.5 Prove deterministic authored build and second import with one comprehensive fixture.

## 4. PPTX projection and lossless lowering

- [x] 4.1 Project trusted imported PPTX deck/page state and every visible object into typed PPJ or an opaque descriptor.
- [x] 4.2 Vendor the read-only source into `<deck>.assets/source/<sha256>.pptx` and bind only relative URI plus SHA-256 in PPJ.
- [x] 4.3 Translate existing imported edit capabilities into revision-bound nativeRef descriptors without exposing raw package identities.
- [x] 4.4 Reproject the source at build time and compute a semantic old/new PPJ diff with stable changed-node and mutation-footprint evidence.
- [x] 4.5 Lower supported differences into the existing source-bound Edit Plan and reject unsupported, ambiguous, stale, or cross-object mutations.
- [x] 4.6 Return source bytes exactly for an unchanged projected PPJ and retain every unknown part, relationship, and opaque graph for edited output.

## 5. Embedded authored-program recovery

- [x] 5.1 Define reserved OPC content types and relationships for `/officeKit/program.ppj` and `/officeKit/program-map.json`.
- [x] 5.2 Embed canonical authored PPJ, stable native IDs, asset hashes, and relevant fingerprints during source-free compilation.
- [x] 5.3 Recover exact PPJ and assets when the embedded program/map are present and structurally valid.
- [x] 5.4 Apply the PPJ-authoritative native-drift policy without prompting, merging drift, or overwriting the input artifact.
- [x] 5.5 Fall back to ordinary projected import when the embedded program is absent or unusable.

## 6. Standalone PPJ CLI

- [x] 6.1 Add lazy `officekit ppj` routing and common bounded file/path/output handling without loading the codec on root import.
- [x] 6.2 Implement `ppj import` and `ppj inspect`, including fuzzy multi-result discovery with stable IDs and no implicit mutation.
- [x] 6.3 Implement `ppj check` and deterministic `--fix` formatting/default repair without semantic rewriting.
- [x] 6.4 Implement `ppj build` with mandatory check, distinct output enforcement, build receipts, and no automatic render/review.
- [x] 6.5 Implement separate `ppj render` and `ppj review` commands with bounded page selection and honest evidence labels.
- [x] 6.6 Add concise human output and stable `--json` contracts for all six commands.

## 7. Optional Task and resume integration

- [x] 7.1 Extend the Task artifact model to recognize PPJ revisions and their authored or projected source identity.
- [x] 7.2 Save immutable PPJ revision, receipt, candidate, review, and output bindings only when `--task` is supplied.
- [x] 7.3 Resume the latest valid/reviewed PPJ revision into a fresh context without restoring a JavaScript heap.
- [x] 7.4 Keep legacy `ctx.plan` tasks listable but return an explicit unsupported-schema result on 2.0 resume without migration.

## 8. Skill and capability convergence

- [x] 8.1 Build a capability registry that classifies every stable Presentation API as PPJ state, nativeRef, compiler/helper, inspect/review, or host-only.
- [x] 8.2 Generate `ppj.md` from the JSON Schema and capability registry with typed fields, limits, minimal examples, and errors.
- [x] 8.3 Rewrite the main Presentations Skill as a short PPJ-first router for create, import, edit, continue, review, and delivery.
- [x] 8.4 Consolidate focused references for fonts, shapes, text, charts/tables, media/layers, motion, components/templates, imported native references, scenarios, and review/delivery.
- [x] 8.5 Remove duplicated task routes, conflicting visual rules, legacy JSX/MJS defaults, and examples that teach generic card/container composition.
- [x] 8.6 Add the host-neutral `presentation-skill-maintainer` Skill and a registry/schema/reference consistency gate.

## 9. Template Creator and Evidence Ledger

- [x] 9.1 Extend Presentation Template schema v3 and search results with optional declared `referenceProgram` and `referencePptx` evidence.
- [x] 9.2 Update the Presentation Template Creator to build and verify clean-room reference PPJ/PPTX while publishing them only when rights allow.
- [x] 9.3 Create the original Evidence Ledger PPJ with hypothesis, method tree, timeline, table, line/bar evidence, confidence interval, decision gates, and sources.
- [x] 9.4 Compile, render, review, and package Evidence Ledger with reference PPJ/PPTX, preview, representative examples, hashes, provenance, and license.
- [x] 9.5 Verify the Evidence Ledger diff does not modify Cranberry Evidence or shared-worktree template WIP.

## 10. Real acceptance and legacy removal

- [x] 10.1 Complete one high-quality authored Evidence Ledger deck and record structural plus page-by-page visual review evidence.
- [x] 10.2 In a fresh context, recover its embedded PPJ, continue editing, and preserve stable identity, design grammar, and non-target pages.
- [x] 10.3 Project the complex 算秩未来 PPTX and complete one typed edit plus one nativeRef edit with no-op and non-target package proof.
- [x] 10.4 Repeat the three workflows in a packed clean install without relying on the repository source tree.
- [x] 10.5 Remove public Presentation/MJS/Compose authoring exports, legacy Skill routes, obsolete examples, and incompatible tests only after PPJ acceptance passes.

## 11. Release and final evidence

- [ ] 11.1 Update Help, generated API docs, architecture, coverage, package inventory, licenses, and release notes for PPJ and the 2.0 break.
- [ ] 11.2 Set package and plugin versions to `2.0.0` and regenerate NativeAOT/proto/package evidence required by the repository.
- [ ] 11.3 Run the final complete npm suite, proto check, NativeAOT build/reproducibility, Skill portability/reference sync, package contents, and release gate once.
- [ ] 11.4 Publish the atomic branch normally, integrate through the current main coordination window without force push, and verify remote main identity.
