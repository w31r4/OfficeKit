## 1. PPJ contract

- [x] 1.1 Allow connector models in the bounded `compositing.opacity` semantic validation and verify the focused connector contract accepts numeric and opacity-token values
- [x] 1.2 Document the connector native owner and canonical `stroke.opacity` projection in the PPJ reference and coverage/backlog records

## 2. Native implementation

- [x] 2.1 Multiply authored connector compositing opacity into the existing native line alpha and verify the existing non-normal/clip/isolation rejection paths remain unchanged
- [x] 2.2 Add the focused connector compositing compile, embedded-snapshot-free projection, and effective-alpha round-trip test

## 3. Verification

- [x] 3.1 Run the focused codec test, strict OpenSpec validation, and `git diff --check`; record any environment-dependent gate skip
