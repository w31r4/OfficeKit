## Context

See proposal.md. PptxConnectorCodec.Apply already replaces a recognized preset only when the requested connector type differs; unchanged geometry and its canonical guides remain intact. The source PPJ compiler currently does not allow connectorType in its field diff.

## Goals / Non-Goals

Goals: reuse the existing geometry replacement and expose exact source authority.
Non-goals: obstacle avoidance, arbitrary bend/guide editing, connector-to-connector targeting or source-owned geometry flattening.

## Decisions

- Add setConnectorType with connectorType as its only field. Issue it only for editable modeled connectors.
- On a changed field, require the issued capability and update the native ConnectorType. Existing native replacement chooses straightConnector1, bentConnector3 or curvedConnector3; style, endpoints and binding metadata stay with their existing owners.
- Keep the required enum and do not invent a delete/default transition. Invalid/missing values reject in schema.
- Use one focused fixture to switch among all three types, checking actual prstGeom, fresh projection, source no-op and non-target package bytes. Keep unrelated preview edits outside the publication scope.

## Risks / Trade-offs

- Unknown source geometry could lose custom routing → retain existing reader proof; opaque connectors get no typed edit capability.
- A type change could accidentally reset end/style state → compare endpoint/arrow/binding values and native geometry in the regression.

## Migration Plan

Additive capability without wire or native codec changes. Run focused connector checks and maintain generated docs before publishing; full F-04 and host routing remain unfinished.
