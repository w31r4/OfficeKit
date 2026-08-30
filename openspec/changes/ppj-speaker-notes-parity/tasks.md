## 1. Language and projection

- [ ] 1.1 Add page-level `setNotes` capability vocabulary and typed notes parsing.
- [ ] 1.2 Project supported imported notes without flattening rich text.
- [ ] 1.3 Issue notes capabilities only from native edit/add evidence.

## 2. Source-bound compile

- [ ] 2.1 Lower plain and topology-preserving rich note edits through the
      existing speaker-notes codec.
- [ ] 2.2 Reject deletion, representation conversion, style/topology changes,
      and unsupported add profiles.

## 3. Agent surface

- [ ] 3.1 Regenerate `ppj.md` and update text/review guidance and coverage.

## 4. Lean verification and integration

- [ ] 4.1 Extend the existing comprehensive PPJ contract with one imported rich
      notes edit and one safely addable plain notes case.
- [ ] 4.2 Run the focused PPJ contract, C# build, Skill-maintainer, and strict
      OpenSpec checks.
- [ ] 4.3 Commit atomically and fast-forward main without force pushing.
