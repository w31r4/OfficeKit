## Context

See proposal.md. Native paths retain Width/Height as nonnegative integers; zero maps to omitted native w/h and selects the shape-coordinate default per axis. PPJ currently requires all paths to share the first positive extent.

## Goals / Non-Goals

Represent every bounded native path extent for custom shapes, including heterogeneous and independently default axes. Keep mask/clip overrides and host appearance separate.

## Decisions

Use optional viewport {width,height}, both required, each 0..2147483.647 in existing 1000-native-unit path units. Omission inherits geometry.viewBox width/height; zero is the native default, not inheritance. Command coordinates still subtract the geometry viewBox origin; this field changes extents only. Projection selects a positive geometry.viewBox baseline from the first path where it fits the existing frame grammar, otherwise 1, and emits viewport only for paths differing from that baseline. Redundant overrides normalize to omission; explicit native zero dimensions normalize to default-axis values, not XML attribute-presence identity. Existing source paths authority supports add/change/remove; removal inherits the current geometry baseline. Extend shape projection only, retaining literal common-positive mask profile.

## Risks / Trade-offs

Zero confused with inheritance -> test independent default axes. First-path changes normalize the common baseline -> compare effective native extents and unchanged commands rather than incidental JSON factoring. Huge native extents overflow frame grammar -> choose bounded positive baseline and retain full per-path value. Preview overclaim -> field diagnostic.

## Migration Plan

Additive schema; existing native wire unchanged. Fresh projections recover previously opaque default/mixed viewport shapes.

The native maximum-extent fixture exposed a PPJ nativeLeaf numeric maximum of 1e9 despite native width/height leaf validation already accepting int.MaxValue. Permit up to 2147483647 only for customGeometryPathWidth/Height; retain the existing numeric ceiling for other leaf kinds. Mask projection cannot emit shape-only viewport overrides, so its common extent must fit the existing geometry.viewBox grammar.
