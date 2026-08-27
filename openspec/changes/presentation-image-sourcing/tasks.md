## 1. Planning and dependencies

- [x] 1.1 Validate the OpenSpec proposal, design, and three capability specifications.
- [x] 1.2 Pin `webfetch-core@0.1.5` and `@iconify-json/lucide@1.2.126` and record their licenses.

## 2. Image model and task assets

- [ ] 2.1 Add shared bounded PNG/JPEG/GIF/SVG byte inspection needed by placement, acquisition, and audit.
- [ ] 2.2 Add `FileBlob` support to direct Presentation and Compose image placement without changing serialization or wire format.
- [ ] 2.3 Implement content-addressed task image assets, private receipts, list/resume behavior, and atomic idempotent writes.

## 3. Discovery and secure acquisition

- [ ] 3.1 Implement license normalization, allowlist enforcement, attribution requirements, and task-bound candidate evidence.
- [ ] 3.2 Implement offline deterministic Lucide search and SVG materialization.
- [ ] 3.3 Implement explicit Openverse and Wikimedia provider search through `webfetch-core` with stable candidate ranking and visible provider reports.
- [ ] 3.4 Implement HTTPS-only DNS-pinned downloading with redirect, byte, MIME, magic, pixel, and dimension limits.

## 4. CLI and audit

- [ ] 4.1 Add lazy `officekit image search|add|list|audit` parsing, help, JSON results, and error contracts.
- [ ] 4.2 Implement candidate/local/URL acquisition routes and source-protecting output behavior.
- [ ] 4.3 Implement PPTX media hash audit, attribution obligations, and optional deterministic sources sidecar.

## 5. Presentation workflow and documentation

- [ ] 5.1 Add a concise Presentations Skill route and one progressive `image-sourcing.md` reference.
- [ ] 5.2 Make review-deliver the authority for crop, clarity, repetition, alt text, visible attribution, and source sidecars.
- [ ] 5.3 Update Help/API, coverage, release notes, third-party notices, package inventory, and version metadata for `1.1.0`.

## 6. Lean verification and delivery

- [ ] 6.1 Extend existing tests with one offline provider/rights contract and one secure download boundary sample.
- [ ] 6.2 Extend existing task/CLI/Presentation tests with add-list-resume-audit and FileBlob/Compose round trips.
- [ ] 6.3 Run the Skill portability/reference/package smoke and make one real single-slide image dogfood with render, audit, and second import evidence.
- [ ] 6.4 Run the final npm, package contents, and release gates once; commit each functional slice atomically and push the branch normally.
