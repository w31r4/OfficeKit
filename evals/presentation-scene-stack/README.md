# Presentation scene-stack evidence

This directory records reproducible facts from the `presentation-scene-stack`
change without redistributing the external PPTX corpus.

- `evidence.v1.json` identifies every local reference by source URL, byte size,
  and SHA-256.
- External files remain under ignored `tmp/` directories and are not package
  inputs.
- `noOpExact` compares the complete PPTX bytes returned by an unchanged
  source-bound export.
- A reorder case compares decompressed OPC entries before and after the edit,
  reimports the result, and checks the target's native identity at the requested
  stack index.
- Host-render equality is evidence that the selected non-overlapping reorder
  did not disturb the page. It is not evidence that every possible reorder is
  visually equivalent.

SlidesCarnival references require CC BY 4.0 attribution and are subject to the
publisher's redistribution terms. NASA references retain their record and
author credits. The committed evidence contains no source media or slide bytes.

