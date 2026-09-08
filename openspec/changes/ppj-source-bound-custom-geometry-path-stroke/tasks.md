## 1. Contract and projection

- [x] 1.1 Add the `customGeometryPathStroke` capability contract, schema enum, registry boundary, and generated/reference coverage; verify PPJ JSON and reference-sync checks pass
- [x] 1.2 Project only explicit canonical custom-path stroke booleans with path-order identity and fail closed for omitted, malformed, or unsupported values; verify focused projection assertions

## 2. Source-bound edit

- [x] 2.1 Prove the selected path and canonical requested boolean against source XML, then splice only direct `a:path/@stroke`; verify stale, invalid, and unchanged requests fail closed
- [x] 2.2 Add the authored/imported/source-bound/reprojection regression and verify only the owning SlidePart changes and path content survives

## 3. Validation

- [x] 3.1 Run the focused native test, OpenSpec strict validation, Presentation Skill/reference gates, and `git diff --check`; record the passing results
