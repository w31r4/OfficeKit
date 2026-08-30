## 1. Contract

- [ ] 1.1 Add optional element state and source-state capability to the wire.
- [ ] 1.2 Parse explicit PPJ state without changing schema ID or default output.

## 2. Native state codec

- [ ] 2.1 Implement canonical authored hidden/locked profiles for supported object kinds.
- [ ] 2.2 Recognize exact imported baseline/full-lock profiles and reject partial profiles.
- [ ] 2.3 Apply source-bound state changes only after fresh source proof.

## 3. PPJ projection and Agent surface

- [ ] 3.1 Project recognized state and issue the bounded `setState` capability.
- [ ] 3.2 Update capability registry, generated `ppj.md`, shapes/layers guidance, and coverage.

## 4. Lean verification and integration

- [ ] 4.1 Extend one existing PPJ contract with authored and source-bound state assertions.
- [ ] 4.2 Run the focused PPJ contract, proto, Skill-maintainer, and strict OpenSpec checks.
- [ ] 4.3 Commit the spec, runtime, guidance, and evidence atomically and fast-forward main without force push.

