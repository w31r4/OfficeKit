## 1. Contract

- [x] 1.1 Add the bounded `media.playback.trigger` schema and wire field.
- [x] 1.2 Parse and validate the two trigger values while preserving omission
  compatibility.

## 2. Native lowering and evidence

- [x] 2.1 Lower `onClick` and `onSlideStart` through the existing media timing
  writer.
- [x] 2.2 Add a focused authored XML and embedded recovery test.

## 3. Bookkeeping and gates

- [x] 3.1 Update coverage, F-13/F-10/F-09 backlog wording, and the media Skill
  reference.
- [x] 3.2 Run strict OpenSpec validation, proto/schema/reference, focused
  native tests, and repository diff checks.
