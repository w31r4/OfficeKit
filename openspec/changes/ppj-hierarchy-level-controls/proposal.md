# Change: Add PPJ hierarchy display levels

## Why

PPJ already authors editable treemap and sunburst hierarchies, but it always
renders every declared level. Dense data therefore forces an Agent either to
show unreadable detail or to delete semantic nodes from the program. PPTD's
bounded `levels` control demonstrates the useful intent: keep the full data
hierarchy while choosing how many levels the current slide reveals.

## What changes

- Add optional `data.series[0].levels` to treemap and sunburst series.
- Keep the complete hierarchy in PPJ while rendering only the first N levels.
- Reallocate the visible treemap area or sunburst ring width to the displayed
  levels instead of leaving blank geometry for hidden descendants.
- Preserve exact authored intent through the embedded PPJ snapshot and project
  snapshot-free output honestly as an editable group.

## What does not change

- Hidden descendants are not deleted from PPJ and still participate in parent
  total validation.
- Arbitrary groups are not inferred as treemap or sunburst data.
- No interactive drill-down, runtime expression or per-node visibility DSL is
  introduced.
