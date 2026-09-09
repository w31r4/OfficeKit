## Context

See proposal.md. PPJ currently merges present edges; native NoLeft/Top/Right/BottomInset operations already remove attributes and normalize edit intent. Bounded layout admission excludes those deletion cases.

## Goals / Non-Goals

Complete direct inset lifecycle, including nested and whole-object deletion. Inheritance resolution and exact host text reflow remain outside this increment.

## Decisions

Use one old/new margin-presence helper for text and table paths. Reuse native deletion markers rather than invent a second PPJ operation. Admit valid true markers and extend the simple-style whitelist to margins. Explicit zero and fractional point values retain existing conversion. Keep the rejection regression meaningful by adding a non-removable anchorCenter field after margins becomes removable.

## Risks / Trade-offs

Deleting one edge might remove siblings → compare all other XML attributes and non-target ZIP entries. Removing the last edge can omit margins in fresh projection → restore a new margins object. Table styles can compact → cover structured-text restoration after whole-style deletion. Unknown properties and missing edit authority keep existing rejection guards.
