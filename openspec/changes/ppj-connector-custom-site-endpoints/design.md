## Context

See proposal.md. PPJ frame anchors have their own metadata and coordinate resolver; native target/index bindings use DrawingML stCxn/endCxn. Custom geometry sites now retain numeric/reference values, but the endpoint model has no index and source dependency logic excludes native bindings.

## Goals / Non-Goals

Connect explicit indexes to fully projected custom shape sites, preserving transforms and source identity. Preset site maps, opaque targets, automatic routing and host dragging remain separate increments.

## Decisions

Add a third endpoint alternative {element, connectionSite}, index 0..1023. Resolve custom geometry using current frame dimensions, native formula validation and site coordinates normalized into the existing frame/group transform. Emit native binding instead of frame-anchor metadata for that endpoint. Projection and capability issuance inspect the target shape before enabling custom site semantics. Source dependency updates include semantic geometry changes, frame changes and existing frame native-leaf overlays; geometry native-leaf mutations on a bound target reject explicitly until their final graph can be resolved. Endpoint changes can switch index/target, detach to coordinates or switch to a frame anchor. An attached connector's frame remains derived.

## Risks / Trade-offs

Index lost as auto -> assert native XML and fresh projection. Stale coordinates -> target geometry/frame dependency tests and explicit unsupported leaf rejection. Native preset authority accidentally expanded -> retain existing preset-site regression. Mixed endpoint kinds -> clear each endpoint's previous binding mode independently.

## Migration Plan

Additive schema/model field; no wire changes. Fresh projections expose editable custom binding indexes. Existing preset/opaque source authority remains unchanged.
