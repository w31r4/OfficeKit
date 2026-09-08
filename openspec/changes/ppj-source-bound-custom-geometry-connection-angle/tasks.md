## 1. Native leaf contract

- [x] 1.1 Add the bounded `customGeometryConnectionSiteAngle60000` leaf to PPJ native-leaf projection and the capability registry, issuing it only for direct custom-geometry connection-site angles with canonical bounded integer tokens; verify formula-backed or malformed angles remain undisclosed.
- [x] 1.2 Extend native-leaf normalization and edit-plan proofing so the ordered site index and angle range are revalidated against the exact source before any write; verify stale and invalid leaf edits fail closed.

## 2. Source-bound write path

- [x] 2.1 Add the SlidePart XML token splice and reopen verification for `customGeometryConnectionSiteAngle60000`, replacing only the selected `a:cxn/@ang` value; verify site positions, order, formulas, paths, handles, and non-target parts remain unchanged.
- [x] 2.2 Add a focused authored/imported/source-bound/reprojection regression for one literal angle and run the narrow codec test filter.

## 3. Evidence and gates

- [x] 3.1 Update the F-04 backlog and coverage/reference evidence with the new leaf's strict boundary, then verify `openspec validate <change> --strict`, `git diff --check`, and the PPJ JSON/review smoke gates.
