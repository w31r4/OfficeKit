# Change: PPJ chart axis arrowheads

## Why

PPJ can style native chart axis lines but cannot express the finite DrawingML
arrowheads that analytical diagrams use to show increasing direction. Agents
must currently add a separate decorative connector, which is visually fragile
and no longer bound to the actual chart coordinate system.

## What changes

- Add bounded `axisLineArrow.start/end` state to ordinary PPJ chart axes.
- Compile and project canonical native `a:headEnd` and `a:tailEnd` leaves on
  chart axis lines.
- Permit capability-issued fixed-topology source continuation while leaving
  irregular line graphs source-owned.
- Keep grid lines, series lines, radar spokes and generated vector-chart axes
  outside this profile.

## Impact

- Additive PPJ schema and Office wire fields; wire version remains 2.
- NativeAOT chart axis reader/writer/projector/lowerer changes.
- One existing comprehensive PPJ contract gains authored/imported arrow proof.
