# Scenario: technical engineering

## Audience task

Help a technical audience understand a system, inspect tradeoffs and failure
modes, and make or execute an engineering decision.

## Narrative

Define the operating constraint, show the system boundary and current behavior,
explain the proposed flow, compare alternatives, expose failure and recovery,
then state the migration or ownership decision.

## Density and rhythm

Introduce topology in stages. Alternate architecture views with concrete traces,
metrics, or decision frames. Preserve enough labels for the reader version while
using motion or progressive pages to manage live complexity.

## Visual carriers and archetypes

- context and boundary map;
- request/data/state flow;
- sequence or lifecycle diagram;
- current-versus-target architecture;
- failure path and recovery behavior;
- migration stages with ownership and rollback gates.

## Visual grammar

Use straight edges, stable node categories, explicit direction, visible trust or
ownership boundaries, and consistent connector semantics. Use surface changes
to distinguish domains or state, not for decoration. Pair topology with a trace,
metric, or operational consequence.

## Avoid

Decorative clouds, unexplained arrows, impossible crossings, diagrams without a
reading path, component inventories without boundaries, code screenshots too
small to inspect, and proposals without failure behavior.

## Review questions

- Are system boundaries, ownership, direction, and state explicit?
- Can the audience follow one concrete request or failure end to end?
- Are alternatives evaluated with the same criteria?
- Does the decision include migration, rollback, and operational consequences?
