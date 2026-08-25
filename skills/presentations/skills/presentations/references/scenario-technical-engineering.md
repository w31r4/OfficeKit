# Scenario: technical engineering

## Audience task

Help a technical audience understand a system, inspect tradeoffs and failure
modes, and make or execute an engineering decision. Lock system boundaries,
environment, measured evidence, ownership, assumptions, and terminology.

The deck must distinguish observed behavior from a proposed design. Security,
reliability, and migration claims require a trace, metric, or explicit premise.

## Narrative

Define the operating constraint, show the system boundary and current behavior,
explain the proposed flow, compare alternatives, expose failure and recovery,
then state migration, ownership, or implementation decisions. Follow at least
one concrete request, state transition, or failure end to end.

## Density and rhythm

Introduce topology in stages. Alternate architecture views with traces,
metrics, or decision frames. Preserve enough labels for a reader version while
using staged pages or motion to manage live complexity. Dense diagrams should
be followed by a consequence or decision page, not another inventory.

## Visual carriers and archetypes

- context and trust/ownership boundary map;
- request, data, state, or event flow;
- sequence, lifecycle, or state-transition diagram;
- current-versus-target architecture on consistent semantics;
- failure path, fallback, and recovery behavior;
- metric or trace tied to one topology choice;
- migration stages with owner, gate, rollback, and observability.

Use diagrams for topology and causality, tables for exact contracts or option
criteria, charts for measured behavior, and code or screenshots only when they
remain readable and materially prove the point.

## Visual grammar

Use straight edges, stable node categories, explicit direction, visible trust
or ownership boundaries, and consistent connector semantics. Use surface
changes to distinguish domains or state, not for decoration. Pair topology
with an operational consequence.

Typography must preserve identifiers and distinguish component names, state,
evidence, and annotation. Line style may encode sync/async, control/data, or
current/proposed, but every distinction needs a visible legend or direct label.
Spacing should reduce crossings before decorative routing is attempted.

## Avoid

Decorative clouds, unexplained arrows, impossible crossings, diagrams without a
reading path, component inventories without boundaries, code too small to
inspect, inconsistent semantics between pages, and proposals without failure
or rollback behavior.

## Review questions

- Are system boundaries, ownership, direction, and state explicit?
- Can the audience follow one request or failure end to end?
- Are alternatives evaluated with the same criteria and evidence?
- Does the decision include migration, rollback, and operational consequences?
- If the diagram fails, should the boundary, flow, or carrier be repaired before
  colors, icons, or motion?
