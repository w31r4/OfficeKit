## Context

See proposal.md. Native ApplyAutoFit already removes a canonical choice and rejects duplicate/noncanonical choices. Native percentage delete markers already remove individual attributes. PPJ present-only merge and bounded admission prevent those operations.

## Goals / Non-Goals

Complete direct choice and percentage lifecycle without changing text geometry or resolving inherited host reflow.

## Decisions

Share old/new AutoFit presence inference across text and table paths. Deleting autoFit with its dependent profile selects NoAutoFitMode and clears the profile. With shrink-text retained, missing previously projected percentage fields select their existing deletion markers. Retaining normalAutoFit without shrink-text remains invalid; its empty object remains invalid under the current schema. Add autoFit/normalAutoFit to guarded whole-style removal. Keep existing native canonical markup checks.

## Risks / Trade-offs

Explicit none could be mistaken for absence → inspect actual noAutofit versus no child. Removing a percentage could remove shrink-text → compare only the requested native attribute and retain normAutofit. Mode changes could retain stale percentages → test switching from a populated profile. Source topology/authority may be unsupported → retain existing guards and related tests.
