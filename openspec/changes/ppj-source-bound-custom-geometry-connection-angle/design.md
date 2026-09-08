## Context

The custom-geometry reader already models ordered `cxnLst/cxn` entries,
including literal or formula-backed angle and position values. Native leaves
already carry an ordered index, source hash, expected value, and edit-plan
binding. See the proposal and the new capability spec for the externally
visible boundary.

## Goals / Non-Goals

**Goals:**

- Reuse the existing native-leaf/edit-plan contract for one direct `ang` scalar.
- Keep the connection-site list and every position/reference/topology field fixed.
- Prove stale source and invalid numeric values before XML mutation.

**Non-Goals:**

- Evaluating guide formulas or exposing x/y position edits.
- Adding/removing/reordering sites or editing adjustment handles, paths, or guide formulas.
- Claiming complete DrawingML custom-geometry connection semantics.

## Decisions

- Use `customGeometryConnectionSiteAngle60000` as a numeric native leaf. The
  name records the native unit and avoids pretending that the field is a
  degree float needing rounding.
- Bind by the direct `cxnLst` ordered index. Require one direct connection-site
  list, canonical `ang`/`pos` structure, and unique list ownership; permit
  other sites to keep references or remain non-literal while only a literal
  target receives a leaf.
- Token-splice only the selected `ang` attribute. The edit plan will locate
  `custGeom/cxnLst/cxn[index]`, verify the expected raw token, and replace its
  value without serializing the surrounding custom geometry.
- Keep the existing protocol version. The generic native leaf value/index
  already carries this scalar, so no protobuf field is necessary.

## Risks / Trade-offs

- [Risk] A source may contain a valid-looking connection-site list with
  extension or duplicate structural content. → [Mitigation] Require the
  direct list/site/position proof and fail closed on unexpected children or
  attributes.
- [Risk] A literal angle can be semantically referenced by host-specific
  behavior that PPJ does not model. → [Mitigation] Preserve every other
  connection-site field and claim only the scalar token, without evaluating or
  rewriting dependent geometry.
