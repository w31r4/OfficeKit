## 1. Freeze and inventory the real benchmark

- [x] 1.1 Add a versioned manifest for the three external PPTX hashes, package inventories, native structure counts, editable node indices, and at least six declared edit targets.
- [x] 1.2 Record native renders and current Kimi/HTML/PPTD comparison evidence without making competitor availability an acceptance dependency.
- [x] 1.3 Add a bounded benchmark runner and independent package oracle that never packages or modifies the source assets.

## 2. Establish exact no-op and text-leaf editing

- [x] 2.1 Preserve unchanged imported shape/group wire projections and prove the complete presentation is unchanged before returning exact source bytes.
- [x] 2.2 Add the additive wire-v2 `APPLY_PPTX_EDIT_PLAN` request/result messages and deterministic JavaScript text-leaf compiler.
- [x] 2.3 Implement the C# read-only structural proof and UTF-8 XML token splice with source, part, element, semantic, leaf, and old-value preconditions.
- [x] 2.4 Prove only declared OPC parts change, masked target XML is byte-identical, output reimports, and stale/ambiguous/scope-expanding plans fail closed.
- [x] 2.5 Validate the “算秩未来” title and blue-gray subtitle edits while unrelated unknown geometry remains preserved and non-blocking.

## 3. Persist compiler evidence in durable tasks

- [x] 3.1 Add private `operations/` task storage and a bounded immutable Edit Plan record linked to source/output revision hashes.
- [x] 3.2 Persist an operation record only after a matching passed review advances task HEAD; reject corrupt, stale, escaping, or oversized records.
- [x] 3.3 Return prior operation evidence on resume while continuing to restore reviewed bytes and rebuild model/node indices rather than JavaScript heap state.

## 4. Add controlled native-leaf editing

- [x] 4.1 Add `presentation.inspect({ includeNativeLeaves: true })` with revision-bound safe text leaf IDs and expected hashes.
- [x] 4.2 Add `presentation.editNativeLeaf(targetId, leafId, { expectedHash, value })` and reject raw XML, XPath, part paths, arbitrary attributes, identities, relationships, namespaces, and topology edits.
- [x] 4.3 Extend the capability-issued leaf registry to codec-proven color and local geometry scalars with positive and stale/cross-revision negative tests.
- [x] 4.4 Expose source-bound leaf IDs for repeated component occurrences and add an atomic `presentation.editComponentOccurrence` batch that validates every issued leaf before mutation.

## 5. Complete the deterministic real edit matrix

- [x] 5.1 Run ordinary and grouped text replacement three times from clean sources with identical output and footprint hashes.
- [x] 5.2 Run textbox move/resize, image placement/crop or same-format replacement, and supported chart title/data edits three times from clean sources.
- [x] 5.3 Select and validate a real SmartArt text or equivalent native leaf without relationship or topology drift.
- [x] 5.4 Validate multi-round REPL edit, review, commit, resume, node-index rebuild, continued edit, and publication.

## 6. Independent acceptance and release gates

- [x] 6.1 Prove all three no-ops byte-identical; all non-target parts exact; all target XML masked exact; all advanced structure counts and relationships stable.
- [ ] 6.2 Prove non-target pages pixel-identical with the unified renderer and complete Windows desktop PowerPoint open/browse/save-copy acceptance without repair prompts.
- [x] 6.3 Run three independent fresh-workspace Agent tasks that inspect/resolve,
  use typed or controlled native edits, reimport, package-diff, compare
  inherited findings, and preserve the source. Visual review is explicitly
  unavailable on the macOS host and is not claimed as completed.
- [x] 6.4 Repeat the accepted workflow from a packed clean installation. The
  self-contained distribution smoke runs without a local OfficeKit install and
  proves source-bound inspection, reuse, reimport, and source protection.
- [x] 6.5 Pass fast gates, full `npm test`, OfficeKit C#, protobuf checks, reproducible WASM, package/release gates, and hosted CI. Hosted slow run 32547495729 passed all segments and the required package/.NET/release checks; publication still needs the release owner's npm credentials.
- [ ] 6.6 Publish the final report with source hashes, OfficeKit evidence, available controls, limitations, and every completion-gate result before marking the persistent Goal complete.
