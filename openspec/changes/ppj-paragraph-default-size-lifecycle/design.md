## Context

See proposal.md. Default-run writing uses hundredth-point integers and Math.Round; source paragraph mutations currently authorize alignment, tabs, bold and italic only.

## Goals / Non-Goals

Complete size presence for ordinary text/shape paragraphs. Placeholder inheritance, table paragraph profiles and other default fields remain separate.

## Decisions

Extend the existing field-specific mask, authority and default-style clone path. Normalize requested size to native hundredths before semantic comparison, using existing nearest-even rounding and the 1..768pt range; reject out-of-range values. Do not narrow all textStyle.size consumers through a global schema edit.

Extend the default-run scalar patch to sz so retained effects and unknown XML do not get rebuilt. Reuse the lifecycle fixture for number and boolean values, checking other parts and non-target paragraph/run XML unchanged.

## Risks / Trade-offs

Sub-hundredth input loses precision → document and assert fresh projection of the rounded value. Broad default-style rewrites can alter unrelated effects → patch only changed scalar attributes. The separately reproduced whole-default-style baseline failure stays explicitly excluded from the related test selection.
