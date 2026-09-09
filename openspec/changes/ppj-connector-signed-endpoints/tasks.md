## 1. Signed coordinates

- [x] 1.1 Add bounded signed endpoint handling in resolver/native reader/writer and connector placement proof. Verify the existing rotated cross-group fixture now succeeds with exact signed coordinates and native validation.
- [x] 1.2 Verify fresh-source signed literal endpoint and frame edits, no-op bytes and non-target ZIP preservation; reject out-of-range authored/native frames without overflow or loss of source bytes.

## 2. Discoverability and publication

- [x] 2.1 Update schema description, Help, capability entry, focused shapes reference and backlog/coverage. Run focused connector regressions, affected preview tests, generated reference/matrix and strict OpenSpec checks before atomic publication.

Evidence (2026-09-10, isolated base 066a6924): 65/65 native Connector and PpjPreviewAuthoredSceneTests pass, zero skipped. The rotated/flip cross-group case now exports signed local endpoint (1.25,-18.75), preserves exact target/anchor after removing embedded PPJ and retains source no-op bytes. Literal (-20,-10)->(40,30) inspects native off values, edits from.x=-40 and separately frame.x=-40 from the original projection, and verifies fresh endpoints plus SlidePart-only differences. Native out-of-range offset remains opaque and byte-identical on native export; PPJ projection explicitly rejects its oversized frame. Extreme endpoints and oversized extents reject without overflowing.

Two test corrections were evidence-driven: the malformed-frame case must distinguish native preservation from PPJ frame limits; an attempted leftEmu leaf assertion was removed after confirming DescribeConnector never issued frame leaves, with the required existing setFrame round trip retained. No new connector leaf API is claimed.

Maintainer and generated matrix checks, scene view/SVG, preview capability coverage, Skill portability (255 files), root reference sync (333 files), whitespace and strict OpenSpec checks pass. No proto field changed. Source-library evidence only; no NativeAOT rebuild, full-repository suite, production preview switch or PowerPoint host acceptance. Full F-04 and the P0/P1 backlog remain open.
